from __future__ import annotations

import argparse
import base64
import calendar
import json
import os
import queue
import re
import socket
import subprocess
import sys
import threading
import time
import traceback
import tkinter as tk
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from playwright.sync_api import Error as PlaywrightError, Page, TimeoutError as PlaywrightTimeoutError, sync_playwright


MODULES = {
    "现场销售": "bfc_sale_local_desk",
    "库房销售": "bfc_sale_desk",
    "拆解料销售": "bfc_sale_cjl_desk",
}

FORM_MODULES = {
    "综合付款单": "综合付款单",
    "车辆付款单": "车辆付款单",
    "采购单": "采购单",
}

ERP_URL = "https://erp.bfcgj.com/module.jsp?module=desk_main"
DEBUG_PORT = 9222


def application_dir() -> Path:
    """Return the script folder, or the executable folder when packaged."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


class LoginRequired(RuntimeError):
    """Normal first-run state: the dedicated ERP browser needs a login."""


@dataclass
class SaleRow:
    sale_type: str
    sale_no: str
    sale_date: str
    payment_date: str
    material: str
    recovery_no: str
    total_amount: str
    tax: str
    total_weight: str
    remark: str


def clean(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip()


def safe_filename(value: str) -> str:
    return re.sub(r'[<>:"/\\|?*]', "_", clean(value))


def open_output_directory(path: Path) -> None:
    """Open a generated output directory in Windows Explorer."""
    directory = path if path.is_dir() else path.parent
    if directory.exists():
        os.startfile(str(directory))


def start_debug_edge() -> None:
    """Open a visible Edge window with a persistent automation profile."""
    edge_paths = (
        Path(r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"),
        Path(r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"),
    )
    edge = next((item for item in edge_paths if item.exists()), None)
    if edge is None:
        raise RuntimeError("未找到 Microsoft Edge，请安装 Edge 后重试")
    profile = application_dir() / ".facas-edge-profile"
    profile.mkdir(exist_ok=True)
    subprocess.Popen([
        str(edge),
        f"--remote-debugging-port={DEBUG_PORT}",
        f"--user-data-dir={profile}",
        "--new-window",
        ERP_URL,
    ])


def debug_port_available() -> bool:
    try:
        with socket.create_connection(("127.0.0.1", DEBUG_PORT), timeout=0.5):
            return True
    except OSError:
        return False


def live_erp_page(browser):
    pages = [
        page for context in browser.contexts for page in context.pages
        if not page.is_closed() and "erp.bfcgj.com" in page.url
    ]
    if not pages:
        raise RuntimeError("未找到仍然打开的 ERP 页面，请保持专用 Edge 和 ERP 页面打开")
    return pages[-1]


def save_page_pdf_fallback(page: Page, path: Path) -> None:
    session = page.context.new_cdp_session(page)
    result = session.send("Page.printToPDF", {
        "printBackground": True,
        "preferCSSPageSize": True,
        "paperWidth": 8.27,
        "paperHeight": 11.69,
        "marginTop": 0.2,
        "marginBottom": 0.2,
        "marginLeft": 0.2,
        "marginRight": 0.2,
    })
    path.write_bytes(base64.b64decode(result["data"]))


def find_print_markup(value):
    """Find HTML-like print markup in a PrintBill response recursively."""
    if isinstance(value, str):
        text = value.strip()
        if "<html" in text.lower() or "<body" in text.lower() or "lodop" in text.lower():
            return text
    if isinstance(value, dict):
        # PrintBill returns the template under anys.<module>.billprint.html;
        # it is a document fragment rather than a complete HTML document.
        for key, item in value.items():
            if key.lower() == "billprint.html" and isinstance(item, str):
                return item
            found = find_print_markup(item)
            if found:
                return found
    if isinstance(value, list):
        for item in value:
            found = find_print_markup(item)
            if found:
                return found
    return None


def page_print_markup(page: Page, expected_values: list[str] | None = None) -> str | None:
    """Read the bill HTML populated by plugInPrint.html after PrintBill."""
    # plugInPrint.html can be a hidden iframe or a separate context page.
    expected = [clean(value) for value in (expected_values or []) if clean(value) and len(clean(value)) >= 2]
    deadline = time.monotonic() + 6
    while time.monotonic() < deadline:
        best_markup = None
        best_score = -1
        for candidate_page in page.context.pages:
            for frame in candidate_page.frames:
                try:
                    bill = frame.locator("#bill1")
                    if not bill.count():
                        continue
                    body = bill.inner_html().strip()
                    if not body:
                        continue
                    text = clean(bill.inner_text())
                    score = sum(1 for value in expected if value in body or value in text)
                    if score > best_score:
                        style = frame.locator("#billStyle")
                        css = style.inner_text() if style.count() else ""
                        best_markup = f"<html><head><style>{css}</style></head><body>{body}</body></html>"
                        best_score = score
                except Exception:
                    continue
        # Never accept a zero-match frame: it may be the hidden print frame
        # left behind by the previous bill.
        if best_markup is not None and expected and best_score > 0:
            return best_markup
        time.sleep(0.2)
    return None


def print_payload_values(value) -> list[str]:
    values = []
    if isinstance(value, dict):
        for item in value.values():
            values.extend(print_payload_values(item))
    elif isinstance(value, list):
        for item in value:
            values.extend(print_payload_values(item))
    elif isinstance(value, (str, int, float)):
        text = clean(str(value))
        if len(text) >= 2 and not text.startswith("data:"):
            values.append(text)
    return values


def dismiss_erp_print_error(page: Page, log) -> None:
    """Close the ERP's DOM error dialog raised when LODOP is unavailable."""
    try:
        dialog = page.locator(".yc-view-message:visible").last
        if not dialog.count():
            return
        message = clean(dialog.inner_text())
        if message:
            log(f"ERP 打印插件提示，已关闭: {message}")
        confirm = dialog.get_by_role("button", name="确认")
        if confirm.count() and confirm.is_visible():
            confirm.click(timeout=2000)
        else:
            dialog.locator("button").last.click(timeout=2000)
    except (PlaywrightError, PlaywrightTimeoutError):
        pass


def ensure_clodop(page: Page, log) -> bool:
    """Load the local C-Lodop bridge when ERP did not include it in the page."""
    try:
        ready = page.evaluate("""() => !!(window.LODOP && typeof window.LODOP.SET_LICENSES === 'function')""")
        if ready:
            return True
        page.evaluate("""() => new Promise((resolve, reject) => {
            if (window.LODOP && typeof window.LODOP.SET_LICENSES === 'function') { resolve(true); return; }
            const script = document.createElement('script');
            script.src = 'http://localhost:8000/CLodopfuncs.js';
            script.onload = () => resolve(true);
            script.onerror = () => reject(new Error('无法加载本机 C-Lodop 脚本'));
            document.head.appendChild(script);
        })""")
        page.wait_for_function("() => !!(window.LODOP && typeof window.LODOP.SET_LICENSES === 'function')", timeout=5000)
        log("已加载本机 C-Lodop 打印组件")
        return True
    except Exception as exc:
        log(f"C-Lodop 未就绪: {exc}")
        return False


def save_printbill_pdf(page: Page, path: Path, log, expected_values: list[str] | None = None) -> bool:
    """Capture the ERP PrintBill response and render returned markup."""
    # Do not click "单据打印": that action invokes the native Windows print
    # dialog. Preview is the non-destructive route used for PDF extraction.
    print_button = page.get_by_role("button", name="打印预览").last
    if not print_button.count() or not print_button.is_visible():
        return False
    ensure_clodop(page, log)
    try:
        with page.expect_response(lambda response: "_name=PrintBill" in response.url, timeout=10000) as response_info:
            print_button.click(timeout=5000, no_wait_after=True)
        response = response_info.value
        if not page.is_closed():
            page.wait_for_timeout(500)
            dismiss_erp_print_error(page, log)
        body = response.text()
        try:
            payload = response.json()
        except Exception:
            payload = body
        payload_values = print_payload_values(payload)
        payload_business_values = sorted({
            value for value in payload_values
            if re.search(r"\$[A-Z0-9]{6,}", value)
        }, key=len, reverse=True)
        # ERP may return a fixed template without the internal bill number.
        # In that case validate the rendered frame against identifiers from
        # the detail page that is currently open.
        business_values = []
        for value in [*(expected_values or []), *payload_business_values]:
            value = clean(str(value or ""))
            if value and len(value) >= 2 and value not in business_values:
                business_values.append(value)
        if not payload_business_values:
            log("PrintBill 响应未包含内部单号，改用当前单据页面标识校验")
        # The response often contains only the fixed template. Prefer the
        # fully populated #bill1 rendered by the ERP print iframe.
        markup = page_print_markup(page, business_values)
        if not markup:
            response_markup = find_print_markup(payload)
            # A static template without the current bill number is not enough:
            # rendering it would create a valid-looking but stale/empty PDF.
            if response_markup and business_values and any(
                    value in response_markup for value in business_values):
                markup = response_markup
        if not markup:
            debug_path = path.with_suffix(".print-response.txt")
            debug_path.write_text(body, encoding="utf-8")
            html_path = path.with_suffix(".print-page.html")
            html_path.write_text(page.content(), encoding="utf-8")
            log(f"PrintBill 已响应，但未发现可渲染 HTML: {body[:200]}")
            log(f"已保存打印响应和打印页面供排查: {debug_path.name}, {html_path.name}")
            return False
        preview = page.context.new_page()
        try:
            preview.set_content(markup, wait_until="networkidle")
            save_page_pdf_fallback(preview, path)
        finally:
            preview.close()
        return path.exists() and path.stat().st_size > 0
    except Exception as exc:
        log(f"捕获 PrintBill 失败: {exc}")
        return False


def data_grid(page: Page, expected_markers: tuple[str, ...] = ()):
    """Return the visible ERP list grid matching its business headers."""
    grids = page.locator("table.yc-view-grid-table")
    header_markers = ("销售单号", "票据号码", "单据日期", "客户名称", "对方单位")
    candidates: list[tuple[int, object]] = []
    for i in range(grids.count()):
        grid = grids.nth(i)
        try:
            if not grid.is_visible():
                continue
            text = clean(grid.inner_text())
            marker_count = sum(marker in text for marker in header_markers)
            expected_count = sum(marker in text for marker in expected_markers)
            record_count = len(re.findall(r"\$[A-Z]{2}\d{6,}|20\d{2}-\d{2}-\d{2}", text))
            data_rows = grid.locator("tbody tr").count() or grid.locator("tr").count()
            has_action = grid.get_by_text("查看", exact=True).count() > 0
            subtotal_penalty = 20 if any(label in text for label in ("本页小计", "合计")) else 0
            if data_rows and (marker_count or record_count):
                # ERP separates the header and body into different tables.
                # Real records must outweigh header-marker matches, otherwise
                # an empty header table wins and detail rows cannot be opened.
                score = (record_count * 50 + expected_count * 20
                         + marker_count * 5 + data_rows + int(has_action) * 3
                         - subtotal_penalty)
                candidates.append((score, grid))
        except PlaywrightError:
            continue
    if candidates:
        return max(candidates, key=lambda item: item[0])[1]
    expected_text = "、".join(expected_markers) if expected_markers else "销售单号或票据号码"
    raise RuntimeError(f"未识别到包含{expected_text}的数据表格")


def click_visible_menu_item(page: Page, name: str) -> None:
    # The ERP may render the navigation tree inside an iframe and may wrap the
    # label in a non-clickable span. Search every frame and click its nearest
    # clickable ancestor, preferring the leftmost visible match.
    visible_candidates = []
    for frame in page.frames:
        candidates = frame.get_by_text(name, exact=True)
        for index in range(candidates.count()):
            candidate = candidates.nth(index)
            try:
                if not candidate.is_visible():
                    continue
                clickable = candidate.locator("xpath=ancestor-or-self::*[self::button or self::a or @role='menuitem' or contains(@class,'menu')][1]")
                target = clickable if clickable.count() else candidate
                box = target.bounding_box()
                if box:
                    visible_candidates.append((box["x"], box["y"], target))
            except PlaywrightError:
                continue
    for _, _, candidate in sorted(visible_candidates, key=lambda item: (item[0], item[1])):
        try:
            candidate.scroll_into_view_if_needed(timeout=3000)
            candidate.click(timeout=10000, force=True)
            page.wait_for_timeout(500)
            return
        except PlaywrightError:
            continue
    raise RuntimeError(f"未找到可点击的菜单项: {name}")


def expand_menu_section(page: Page, name: str, child_name: str | None = None) -> None:
    """Expand a first-level ERP menu without collapsing an already-open one."""
    for frame in page.frames:
        labels = frame.get_by_text(name, exact=True)
        for index in range(labels.count()):
            label = labels.nth(index)
            try:
                if not label.is_visible():
                    continue
                target = label.locator("xpath=ancestor-or-self::*[self::button or self::a or @role='treeitem' or @role='menuitem' or contains(@class,'menu')][1]")
                if not target.count():
                    target = label
                if child_name:
                    child = frame.get_by_text(child_name, exact=True)
                    if any(child.nth(i).is_visible() for i in range(child.count())):
                        return
                expanded = target.get_attribute("aria-expanded")
                classes = (target.get_attribute("class") or "").lower()
                # aria-expanded is authoritative when provided. For older ERP
                # markup, infer collapsed state from common class names.
                collapsed = expanded == "false" or any(token in classes for token in ("collapsed", "closed", "fold"))
                if expanded is None and not collapsed:
                    target.click(timeout=8000, force=True)
                    page.wait_for_timeout(400)
                    return
                if collapsed:
                    target.click(timeout=8000, force=True)
                    page.wait_for_timeout(400)
                return
            except PlaywrightError:
                continue
    raise RuntimeError(f"未找到一级菜单或无法展开: {name}")


def apply_date_filters(filter_table, start: str, end: str, before_search=None) -> bool:
    """Apply the first/last date fields without assuming mixed input indexes."""
    checkboxes = filter_table.locator("input[type='checkbox']:visible")
    text_candidates = filter_table.locator("input:not([type='checkbox']):visible")
    editable_inputs = []
    for index in range(text_candidates.count()):
        candidate = text_candidates.nth(index)
        try:
            if candidate.is_editable():
                editable_inputs.append(candidate)
        except PlaywrightError:
            continue
    if checkboxes.count() < 4 or len(editable_inputs) < 2:
        return False
    # Resolve each checkbox by the label in its nearest filter container,
    # rather than by position (the ERP changes order between modules).
    date_controls = []
    for index in range(checkboxes.count()):
        checkbox = checkboxes.nth(index)
        label_text = checkbox.evaluate("""el => {
            let node = el;
            for (let i = 0; i < 6 && node; i++, node = node.parentElement) {
                const text = (node.innerText || '').replace(/\\s+/g, '');
                if (text.includes('销售日期从') || text.includes('销售日期到')) return text;
            }
            return '';
        }""")
        if "销售日期从" in label_text or "销售日期到" in label_text:
            # A filter field is rendered inside one table cell: checkbox,
            # label, and input live in the same td. Keep the lookup within
            # that exact HTML cell so it cannot select 客户名称 or 回收单号.
            container = checkbox.locator("xpath=ancestor::td[1]")
            date_input = container.locator("input:not([type='checkbox']):visible").first
            if date_input.count():
                value = start if "销售日期从" in label_text else end
                date_controls.append((checkbox, date_input, value))
    if len(date_controls) != 2:
        date_controls = []
        for label, value in (("销售日期从", start), ("销售日期到", end)):
            label_locator = filter_table.get_by_text(label, exact=False).filter(visible=True)
            found = False
            for index in range(label_locator.count()):
                label_node = label_locator.nth(index)
                cell = label_node.locator("xpath=ancestor::td[1]")
                checkbox = cell.locator("input[type='checkbox']:visible").first
                date_input = cell.locator("input:not([type='checkbox']):visible").first
                if checkbox.count() and date_input.count():
                    date_controls.append((checkbox, date_input, value))
                    found = True
                    break
            if not found:
                return False
    for index in range(checkboxes.count()):
        checkbox = checkboxes.nth(index)
        if checkbox.is_checked():
            checkbox.uncheck()
    for checkbox, date_input, value in date_controls:
        checkbox.check()
        date_input.fill(value)
        date_input.press("Enter")
    search = filter_table.get_by_role("button", name="搜索")
    if search.count():
        if before_search is not None:
            before_search()
        search.click()
        return True
    return False


def apply_sale_date_filters(filter_table, start: str, end: str, before_search=None) -> bool:
    """Use the stable sales-page checkbox/input pairs from the ERP layout."""
    inputs = filter_table.locator("input:visible")
    if inputs.count() < 8:
        return False
    # Disable every filter first, then enable only sales-date from/to.
    for index in range(inputs.count()):
        control = inputs.nth(index)
        if (control.get_attribute("type") or "").lower() == "checkbox" and control.is_checked():
            control.uncheck()
    for checkbox_index, input_index, value in ((0, 1, start), (6, 7, end)):
        checkbox = inputs.nth(checkbox_index)
        date_input = inputs.nth(input_index)
        if not checkbox.is_checked():
            checkbox.check()
        date_input.fill(value)
        date_input.press("Enter")
    search = filter_table.get_by_role("button", name="搜索")
    if search.count():
        if before_search is not None:
            before_search()
        search.click()
        return True
    return False


def remove_response_listener(page: Page, callback) -> None:
    """Release a response callback even when the ERP page is closing."""
    try:
        page.remove_listener("response", callback)
    except PlaywrightError:
        pass


def find_named_blocks(value, name: str, module: str | None = None) -> list[dict]:
    """Return every named ERP response block, optionally for one module."""
    found: list[dict] = []
    if isinstance(value, dict):
        if (value.get("name") == name and isinstance(value.get("rowsData"), list)
                and (module is None or value.get("_amn") == module)):
            found.append(value)
        for child in value.values():
            found.extend(find_named_blocks(child, name, module))
    elif isinstance(value, list):
        for child in value:
            found.extend(find_named_blocks(child, name, module))
    return found


DETAIL_FIELD_MAP = {
    # Fields confirmed from the ERP's bfc_sale_local_bill data_d response.
    "现场销售": {
        "material": "c14", "recovery_no": "c3", "amount": "c12",
        "tax": "c17", "weight": "c16",
    },
    # Fields confirmed from the ERP's bfc_sale_bill data_d response.
    "库房销售": {
        "material": "c12", "recovery_no": None, "amount": "c10",
        "tax": "c15", "weight": "c26",
    },
    "拆解料销售": {
        "material": "c14", "recovery_no": None, "amount": "c12",
        "tax": "c17", "weight": "c16",
    },
}


def numeric_value(value) -> float | None:
    """Convert ERP numeric fields while treating blanks as absent values."""
    if value is None or clean(str(value)) == "":
        return None
    try:
        return float(str(value).replace(",", ""))
    except (TypeError, ValueError):
        return None


def format_amount(value: float | None) -> str:
    return f"{value:,.2f}" if value is not None else ""


def erp_date(value) -> str:
    if isinstance(value, (int, float)):
        return datetime.fromtimestamp(value / 1000).strftime("%Y-%m-%d")
    return clean(str(value or ""))


SALE_LIST_FIELD_MAP = {
    "现场销售": {
        "sale_no": "c0", "sale_date": "c4", "payment_date": "c9",
        "customer": "c3", "amount": "c5", "weight": "c16",
    },
    "库房销售": {
        "sale_no": "c0", "sale_date": "c4", "payment_date": "c5",
        "customer": "c3", "amount": "c6", "weight": "c15",
    },
    "拆解料销售": {
        "sale_no": "c0", "sale_date": "c4", "payment_date": "c9",
        "customer": "c3", "amount": "c5", "weight": "c15",
    },
}


def sale_list_row(sale_type: str, row: dict) -> tuple[str, str, str, str, str, str]:
    """Map the ERP data_list response for the selected sales desk."""
    fields = SALE_LIST_FIELD_MAP[sale_type]
    return (
        clean(str(row.get(fields["sale_no"]) or "")),
        erp_date(row.get(fields["sale_date"])),
        erp_date(row.get(fields["payment_date"])),
        clean(str(row.get(fields["customer"]) or "")),
        format_amount(numeric_value(row.get(fields["amount"]))),
        format_amount(numeric_value(row.get(fields["weight"]))),
    )


def click_sale_detail(page: Page, sale_no: str) -> None:
    """Open a visible list row by the API-provided bill number."""
    matches = page.get_by_text(sale_no, exact=True)
    for index in range(matches.count()):
        bill_cell = matches.nth(index)
        try:
            if not bill_cell.is_visible():
                continue
            row = bill_cell.locator("xpath=ancestor::tr[1]")
            view = row.get_by_text("查看", exact=True)
            if view.count() and view.is_visible():
                view.click()
                return
        except PlaywrightError:
            continue
    raise RuntimeError(f"页面中未找到销售单 {sale_no} 的查看操作")


def detail_api_summary(sale_type: str, rows: list[dict]) -> tuple[str, str, str, str, str]:
    """Read the first item and aggregate numeric detail fields from data_d."""
    fields = DETAIL_FIELD_MAP[sale_type]
    material = ""
    recovery_no = ""
    amount_total = tax_total = weight_total = 0.0
    amount_seen = tax_seen = weight_seen = False
    for row in rows:
        if not material:
            material = clean(str(row.get(fields["material"]) or ""))
        if not recovery_no and fields["recovery_no"]:
            recovery_no = clean(str(row.get(fields["recovery_no"]) or ""))
        amount = numeric_value(row.get(fields["amount"]))
        tax = numeric_value(row.get(fields["tax"]))
        weight = numeric_value(row.get(fields["weight"]))
        if amount is not None:
            amount_total += amount
            amount_seen = True
        if tax is not None:
            tax_total += tax
            tax_seen = True
        if weight is not None:
            weight_total += weight
            weight_seen = True
    return (
        material,
        recovery_no,
        format_amount(amount_total if amount_seen else None),
        format_amount(tax_total if tax_seen else None),
        format_amount(weight_total if weight_seen else None),
    )


def sale_master_api_summary(rows: list[dict]) -> tuple[str, str, str]:
    """Read totals, tax total, and remark from a sales data_m response."""
    if not rows:
        return "", "", ""
    row = rows[0]
    return (
        format_amount(numeric_value(row.get("c13"))),
        format_amount(numeric_value(row.get("c14"))),
        clean(str(row.get("c9") or "")),
    )


def voucher_value(row: dict, header: str) -> str:
    header = clean(header)
    key_map = {
        "票据号码": "c0", "单据号码": "c0", "编号": "c0",
        "对方单位": "c9", "供应商": "c9", "客户": "c9",
        "单据日期": "c3", "完成日期": "c4", "日期": "c3",
        "金额": "c6", "结算方式": "c5", "制单": "c7",
    }
    key = next((source for label, source in key_map.items() if label in header), None)
    value = row.get(key, "") if key else ""
    if key in ("c3", "c4") and isinstance(value, (int, float)):
        return datetime.fromtimestamp(value / 1000).strftime("%Y-%m-%d")
    if value is None:
        return ""
    return clean(str(value).replace("\u00a0", " "))


def scrape_module(page: Page, sale_type: str, start: str, end: str, log,
                  pdf_dir: Path | None = None, generated_pdfs: set[Path] | None = None) -> list[SaleRow]:
    if page.is_closed():
        raise RuntimeError("ERP 页面已关闭，请重新登录专用 Edge")
    log(f"{sale_type}: 正在打开销售列表")
    expand_menu_section(page, "销售管理", sale_type)
    click_visible_menu_item(page, sale_type)
    page.wait_for_timeout(500)
    data_grid(page, ("销售单号",)).wait_for(state="visible", timeout=15000)

    sale_page_rows: dict[int, list[dict]] = {}
    sale_page_batches: list[list[dict]] = []
    sale_page_signatures: set[tuple[str, ...]] = set()
    sale_total_pages = 0
    sale_total_rows = 0
    sale_page_size = 0
    capture_sale_enabled = False
    sale_response_seen = False

    def capture_sale_response(response):
        nonlocal sale_total_pages, sale_total_rows, sale_page_size, sale_response_seen
        if not capture_sale_enabled:
            return
        if "erp.bfcgj.com" not in response.url:
            return
        try:
            payload = response.json()
        except Exception:
            return
        for block in find_named_blocks(payload, "data_list", MODULES[sale_type]):
            sale_response_seen = True
            rows = list(block["rowsData"])
            signature = tuple(clean(str(row.get("c0") or "")) for row in rows)
            if signature and signature not in sale_page_signatures:
                sale_page_signatures.add(signature)
                sale_page_batches.append(rows)
            page_index = int(block.get("page") or 0)
            sale_page_rows[page_index] = rows
            sale_total_pages = max(sale_total_pages, int(block.get("pages") or 0))
            sale_total_rows = max(sale_total_rows, int(block.get("totalRows") or 0))
            sale_page_size = max(sale_page_size, int(block.get("pageRows") or len(rows)))

    def begin_sale_capture():
        nonlocal capture_sale_enabled, sale_response_seen, sale_total_pages, sale_total_rows, sale_page_size
        sale_page_rows.clear()
        sale_page_batches.clear()
        sale_page_signatures.clear()
        sale_total_pages = sale_total_rows = sale_page_size = 0
        sale_response_seen = False
        capture_sale_enabled = True

    page.on("response", capture_sale_response)
    output: list[SaleRow] = []
    detail_failures: list[str] = []
    listed_count = skipped_payment_count = page_no = 0
    seen_sale_nos: set[str] = set()
    try:
        filter_tables = page.locator("table.yc-view-free-table:visible")
        filter_table = filter_tables.first if filter_tables.count() else page.locator("table.yc-view-free-table").first
        # Ignore any delayed default-list response emitted while the desk was
        # opening. Only responses triggered by the requested filter/pagination
        # are allowed into the page buffers.
        sale_page_rows.clear()
        sale_page_batches.clear()
        sale_page_signatures.clear()
        sale_total_pages = sale_total_rows = sale_page_size = 0
        sale_response_seen = False
        capture_sale_enabled = False
        if not apply_sale_date_filters(filter_table, start, end, begin_sale_capture):
            raise RuntimeError(f"{sale_type}日期筛选未成功执行，已停止导出")
        log(f"{sale_type}: 已提交销售日期筛选 {start} 至 {end}")
        deadline = time.monotonic() + 10
        while not sale_response_seen and time.monotonic() < deadline:
            page.wait_for_timeout(200)
        if not sale_response_seen:
            raise RuntimeError(f"未捕获{sale_type} data_list 响应，已停止导出")
        if not sale_page_batches:
            if sale_total_rows:
                raise RuntimeError(f"{sale_type} data_list 返回空数据但 totalRows={sale_total_rows}")
            log(f"{sale_type}: 日期范围内未查询到销售单")
            return output
        sale_page_rows[0] = sale_page_batches[0]
        expected_pages = 0
        if sale_total_rows and sale_page_size:
            expected_pages = (sale_total_rows + sale_page_size - 1) // sale_page_size
        effective_total_pages = max(sale_total_pages, expected_pages)
        if sale_total_pages:
            log(f"{sale_type}: 接口共 {sale_total_pages} 页")
        if expected_pages and expected_pages != sale_total_pages:
            log(f"{sale_type}: 接口共 {sale_total_rows} 条，预计 {expected_pages} 页")

        while True:
            if page.is_closed():
                raise RuntimeError("ERP 页面已关闭，无法继续读取销售列表")
            page_no += 1
            log(f"{sale_type}: 正在读取第 {page_no} 页")
            current_rows = sale_page_rows.get(page_no - 1)
            if current_rows is None and page_no - 1 < len(sale_page_batches):
                current_rows = sale_page_batches[page_no - 1]
            if current_rows is None:
                raise RuntimeError(f"未捕获{sale_type}第 {page_no} 页 data_list 响应")
            for sale_no, sale_date, payment_date, _customer, _amount, _list_weight in (
                    sale_list_row(sale_type, row) for row in current_rows):
                if sale_no in seen_sale_nos:
                    continue
                seen_sale_nos.add(sale_no)
                listed_count += 1
                if not payment_date or not re.search(r"\d{4}-\d{2}-\d{2}", payment_date):
                    skipped_payment_count += 1
                    continue
                material = recovery_no = remark = tax = ""
                amount = weight = ""
                detail_succeeded = False
                close = None
                capture_detail_response = None
                try:
                    log(f"{sale_type}: 正在读取销售单 {sale_no} 详情")
                    detail_response_rows: list[dict] = []
                    master_response_rows: list[dict] = []
                    detail_row_keys: set[tuple] = set()
                    master_row_keys: set[tuple] = set()

                    def response_row_key(row: dict) -> tuple:
                        """Identify repeated ERP rows across duplicate responses."""
                        for field in ("__rowid", "__rownum"):
                            value = row.get(field)
                            if value is not None and clean(str(value)):
                                return field, str(value)
                        return "data", json.dumps(row, ensure_ascii=False, sort_keys=True, default=str)

                    def capture_detail_response(response):
                        if "erp.bfcgj.com" not in response.url:
                            return
                        try:
                            payload = response.json()
                        except Exception:
                            return
                        detail_module = MODULES[sale_type].replace("_desk", "_bill")
                        for block in find_named_blocks(payload, "data_d", detail_module):
                            for row in block["rowsData"]:
                                if clean(str(row.get("c0") or "")) != sale_no:
                                    continue
                                key = response_row_key(row)
                                if key not in detail_row_keys:
                                    detail_row_keys.add(key)
                                    detail_response_rows.append(row)
                        for block in find_named_blocks(payload, "data_m", detail_module):
                            for row in block["rowsData"]:
                                if clean(str(row.get("c0") or "")) != sale_no:
                                    continue
                                key = response_row_key(row)
                                if key not in master_row_keys:
                                    master_row_keys.add(key)
                                    master_response_rows.append(row)

                    page.on("response", capture_detail_response)
                    click_sale_detail(page, sale_no)
                    close = page.get_by_role("button", name="关闭").last
                    close.wait_for(state="visible", timeout=10000)
                    deadline = time.monotonic() + 10
                    while (not detail_response_rows or not master_response_rows) and time.monotonic() < deadline:
                        page.wait_for_timeout(200)
                    if not detail_response_rows or not master_response_rows:
                        raise RuntimeError("未捕获销售详情完整 data_m/data_d 响应")
                    material, recovery_no, calculated_amount, detail_tax, weight = detail_api_summary(sale_type, detail_response_rows)
                    master_amount, master_tax, remark = sale_master_api_summary(master_response_rows)
                    amount = master_amount or calculated_amount
                    tax = master_tax if numeric_value(master_tax) not in (None, 0.0) else detail_tax
                    detail_succeeded = True
                    log(f"{sale_type}: 已读取销售单 {sale_no} 详情")
                    if pdf_dir is not None:
                        pdf_path = pdf_dir / f"{safe_filename(sale_type)}-{safe_filename(sale_date)}-{safe_filename(sale_no)}.pdf"
                        expected_print_values = [sale_no, sale_no.lstrip("$")[-8:], sale_date]
                        if save_printbill_pdf(page, pdf_path, log, expected_print_values):
                            if generated_pdfs is not None:
                                generated_pdfs.add(pdf_path.resolve())
                            log(f"已保存 PDF: {pdf_path.name}")
                        else:
                            log(f"未生成票据 PDF: {pdf_path.name}")
                except (RuntimeError, PlaywrightTimeoutError, PlaywrightError, IndexError, OSError) as exc:
                    detail_failures.append(sale_no)
                    log(f"详情读取失败，已跳过销售单 {sale_no}: {exc}")
                finally:
                    if capture_detail_response is not None:
                        remove_response_listener(page, capture_detail_response)
                    try:
                        if close is not None and not page.is_closed() and close.count() and close.is_visible():
                            close.click(timeout=3000)
                            page.wait_for_timeout(300)
                    except PlaywrightError:
                        log(f"详情关闭失败: {sale_no}")
                if detail_succeeded:
                    output.append(SaleRow(sale_type, sale_no, sale_date, payment_date, material, recovery_no, amount, tax, weight, remark))

            # The API's page count is authoritative. Do not click the next
            # button after the final page, where no new data_list response can
            # arrive and the old timeout misleadingly reports a missing page.
            if effective_total_pages and page_no >= effective_total_pages:
                break
            target_page = str(page_no + 1)
            page_links = page.locator("li.paging-item:visible")
            next_buttons = page.locator('[paging-btn="nextPage"]:visible')
            advance = None
            for link_index in range(page_links.count()):
                candidate = page_links.nth(link_index)
                if clean(candidate.inner_text()) == target_page:
                    advance = candidate
                    break
            if advance is not None:
                pass
            else:
                if not next_buttons.count():
                    break
                advance = next_buttons.last
                if advance.is_disabled() or "disabled" in (advance.get_attribute("class") or ""):
                    break
            previous_batch_count = len(sale_page_batches)
            selected_page = ""
            selected_items = page.locator("li.paging-item.paging-item-selected:visible")
            if selected_items.count():
                selected_page = clean(selected_items.last.inner_text())
            already_selected = selected_page == target_page
            batch_ready = already_selected and len(sale_page_batches) > page_no
            if already_selected and len(sale_page_batches) > page_no:
                sale_page_rows[page_no] = sale_page_batches[page_no]
            elif not already_selected:
                advance.scroll_into_view_if_needed()
                advance.click(timeout=10000, force=True)
            deadline = time.monotonic() + 10
            if not batch_ready:
                while len(sale_page_batches) <= previous_batch_count and time.monotonic() < deadline:
                    page.wait_for_timeout(200)
                if len(sale_page_batches) <= previous_batch_count:
                    raise RuntimeError(f"未捕获{sale_type}第 {page_no + 1} 页 data_list 响应")
                sale_page_rows[page_no] = sale_page_batches[previous_batch_count]
    finally:
        remove_response_listener(page, capture_sale_response)
    failure_note = f"；详情失败 {len(detail_failures)} 张: {'、'.join(detail_failures)}" if detail_failures else ""
    log(f"{sale_type}: 列表销售单 {listed_count} 张，因收款日期为空跳过 {skipped_payment_count} 张，导出 {len(output)} 张{failure_note}")
    return output


def write_sheet(ws, title: str, headers: list[str], rows: list[list[str]], widths: list[int] | None = None) -> None:
    ws.title = title[:31]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")
    for row in rows:
        ws.append(row)
    ws.freeze_panes = "A2"
    if widths is not None:
        for index, width in enumerate(widths, 1):
            ws.column_dimensions[chr(64 + index)].width = width
        return
    for column in ws.columns:
        letter = column[0].column_letter
        max_len = min(30, max([len(str(cell.value or "")) for cell in column] + [10]))
        ws.column_dimensions[letter].width = max_len + 2


def export_report(rows: list[SaleRow], sheets: dict[str, tuple[list[str], list[list[str]]]], path: Path) -> None:
    """Write only data selected in this run, replacing any prior report."""
    wb = Workbook()
    active_sheet = wb.active
    has_sheet = False
    if rows:
        headers = ["销售类型", "销售单号", "销售日期", "收款日期", "物料名称", "回收单号", "总金额", "税金", "总重量", "备注"]
        data = [[row.sale_type, row.sale_no, row.sale_date, row.payment_date, row.material,
                 row.recovery_no, row.total_amount, row.tax, row.total_weight, row.remark] for row in rows]
        write_sheet(active_sheet, "销售数据", headers, data, [12, 20, 14, 14, 22, 22, 14, 14, 14, 30])
        for row in active_sheet.iter_rows(min_row=2):
            row[4].alignment = row[5].alignment = row[9].alignment = Alignment(wrap_text=True, vertical="top")
        has_sheet = True
    for name, (headers, data) in sheets.items():
        ws = wb.create_sheet() if has_sheet else active_sheet
        write_sheet(ws, name, headers, data)
        has_sheet = True
    wb.save(path)


def collect_voucher_response_rows(page: Page, start: str, end: str, log, form_name: str,
                                  open_form) -> list[dict]:
    """Collect vc_bill_list responses and always detach the page listener."""
    response_rows: list[dict] = []
    response_count = 0
    total_pages = 0
    total_rows = 0
    page_size = 0
    capture_enabled = False

    def capture_response(response):
        nonlocal response_count, total_pages, total_rows, page_size
        if not capture_enabled:
            return
        if "erp.bfcgj.com" not in response.url:
            return
        try:
            payload = response.json()
        except Exception:
            return
        blocks = find_named_blocks(payload, "vc_bill_list")
        if blocks:
            for block in blocks:
                rows = list(block["rowsData"])
                response_rows.extend(rows)
                total_pages = max(total_pages, int(block.get("pages") or 0))
                total_rows = max(total_rows, int(block.get("totalRows") or 0))
                page_size = max(page_size, int(block.get("pageRows") or len(rows)))
            if not total_pages and total_rows and page_size:
                total_pages = (total_rows + page_size - 1) // page_size
            response_count += 1

    def wait_for_response(previous_count: int, timeout: float = 10, phase: str = "翻页") -> None:
        deadline = time.monotonic() + timeout
        while response_count <= previous_count and time.monotonic() < deadline:
            page.wait_for_timeout(200)
        if response_count <= previous_count:
            raise RuntimeError(f"{form_name}: {phase}未捕获 vc_bill_list 响应")

    def begin_filtered_capture():
        nonlocal capture_enabled, response_count, total_pages, total_rows, page_size
        response_rows.clear()
        response_count = 0
        total_pages = 0
        total_rows = 0
        page_size = 0
        capture_enabled = True

    page.on("response", capture_response)
    try:
        filter_table = open_form()
        page.wait_for_timeout(300)
        # The form opening request returns the unfiltered default list.
        # Discard it before applying the requested date range so those
        # rows cannot leak into the export.
        if response_rows:
            log(f"{form_name}: 已丢弃默认列表 {len(response_rows)} 条，准备应用日期筛选")
        response_rows.clear()
        response_count = 0
        total_pages = 0
        total_rows = 0
        page_size = 0
        capture_enabled = False
        before_search = response_count
        if apply_date_filters(filter_table, start, end, begin_filtered_capture):
            wait_for_response(before_search, phase="日期筛选")
        else:
            search = filter_table.get_by_role("button", name="搜索")
            if search.count():
                begin_filtered_capture()
                before_search = response_count
                search.click()
                wait_for_response(before_search, phase="搜索")
            else:
                raise RuntimeError(f"{form_name}: 未找到搜索按钮")
        page_no = 1
        while True:
            log(f"{form_name}: 正在读取第 {page_no} 页")
            if total_pages and page_no >= total_pages:
                log(f"{form_name}: 接口共 {total_pages} 页")
                break
            # A stale ERP pager can remain visible after a form switch. If
            # the response did not provide page metadata, do not click it and
            # wait for a response that may belong to another view.
            if not total_pages:
                log(f"{form_name}: 接口未提供可靠分页信息，已停止在第 {page_no} 页")
                break
            next_button = page.locator('button[paging-btn="nextPage"]:visible').last
            if not next_button.count() or next_button.is_disabled() or "disabled" in (next_button.get_attribute("class") or ""):
                break
            before_page = response_count
            next_button.click()
            wait_for_response(before_page, phase="翻页")
            page_no += 1
        # The PDF pass runs after this listener is detached; preserve the
        # server-reported page count for that second pass.
        setattr(page, "_facas_voucher_total_pages", total_pages)
        return response_rows
    finally:
        remove_response_listener(page, capture_response)


def click_voucher_page_and_wait(page: Page, button, form_name: str, timeout: float = 10) -> None:
    """Click a voucher pager control and wait for its vc_bill_list response."""
    response_count = 0

    def capture_response(response):
        nonlocal response_count
        if "erp.bfcgj.com" not in response.url:
            return
        try:
            payload = response.json()
        except Exception:
            return
        if find_named_blocks(payload, "vc_bill_list"):
            response_count += 1

    page.on("response", capture_response)
    try:
        button.click(timeout=10000, force=True)
        deadline = time.monotonic() + timeout
        while response_count == 0 and time.monotonic() < deadline:
            page.wait_for_timeout(200)
        if response_count == 0:
            raise RuntimeError(f"{form_name} 翻页后未捕获 vc_bill_list 响应")
    finally:
        remove_response_listener(page, capture_response)


def scrape_default_form(page: Page, form_name: str, start: str, end: str, log,
                        pdf_dir: Path | None = None, generated_pdfs: set[Path] | None = None) -> tuple[list[str], list[list[str]]]:
    """Read a voucher list using the columns rendered by the ERP itself."""
    # Voucher menus are nested under 财务管理 > 凭证处理 > 凭证探测器.
    category = "采购" if form_name == "采购单" else "车辆"
    menu_path = (
        ("财务管理", "凭证处理"),
        ("凭证处理", "凭证探测器"),
        ("凭证探测器", category),
        (category, form_name),
    )
    for parent_name, child_name in menu_path:
        expand_menu_section(page, parent_name, child_name)
    log(f"{form_name}: 正在打开凭证列表")
    def open_form():
        # The response listener must be installed before opening the form;
        # ERP commonly sends the initial vc_bill_list request immediately.
        click_visible_menu_item(page, form_name)
        page.wait_for_timeout(700)
        table = data_grid(page, ("票据号码", "单据日期"))
        table.wait_for(state="visible", timeout=15000)
        filter_tables = page.locator("table.yc-view-free-table:visible")
        return filter_tables.first if filter_tables.count() else page.locator("table.yc-view-free-table").first

    response_rows = collect_voucher_response_rows(page, start, end, log, form_name, open_form)
    voucher_total_pages = getattr(page, "_facas_voucher_total_pages", 0)
    log(f"{form_name}: 已获取 {len(response_rows)} 条接口记录")

    # Confirmed visible order, mapped to the vc_bill_list response fields.
    headers: list[str] = [
        "票据号码", "单据日期", "对方单位", "结算方式",
        "完成日期", "金额", "制单",
    ]
    rows_out: list[list[str]] = []
    if not response_rows:
        raise RuntimeError("未捕获 vc_bill_list 响应，已停止导出以避免生成错误列数据")
    seen_response = set()
    for row in response_rows:
        bill_no = clean(str(row.get("c0", "")))
        if bill_no and bill_no in seen_response:
            continue
        if bill_no:
            seen_response.add(bill_no)
        values = [voucher_value(row, header) for header in headers]
        if any(values) and values not in rows_out:
            rows_out.append(values)
    if pdf_dir is not None:
        pending = {}
        for row in response_rows:
            bill_no = clean(str(row.get("c0", "")))
            if bill_no:
                pending.setdefault(bill_no, row)
        log(f"{form_name}: 开始批量保存 PDF，共 {len(pending)} 张")

        # List extraction ends on the final page. Return to the beginning and
        # process the visible bills page by page so earlier pages are not lost.
        previous_button = page.locator('button[paging-btn="prevPage"]:visible').last
        selected_page = "1"
        selected_items = page.locator("li.paging-item.paging-item-selected:visible")
        if selected_items.count():
            selected_page = clean(selected_items.last.inner_text())
        if selected_page != "1":
            for _ in range(200):
                if (not previous_button.count() or previous_button.is_disabled()
                        or "disabled" in (previous_button.get_attribute("class") or "")):
                    break
                click_voucher_page_and_wait(page, previous_button, form_name)

        while pending:
            grid = data_grid(page, ("票据号码", "单据日期"))
            for bill_no, row in list(pending.items()):
                log(f"{form_name}: 正在打开单据 {bill_no} 详情")
                # The API uses the internal number ($AF...), while the grid
                # normally renders only c1 (for example 00000080). Match
                # either representation instead of assuming both are shown.
                short_bill_no = clean(str(row.get("c1") or ""))
                bill_tokens = [token for token in (bill_no, short_bill_no) if token]
                matching_rows = []
                for row_index in range(grid.locator("tr").count()):
                    candidate_row = grid.locator("tr").nth(row_index)
                    try:
                        if not candidate_row.is_visible():
                            continue
                        row_text = clean(candidate_row.inner_text())
                        if any(token in row_text for token in bill_tokens):
                            matching_rows.append((candidate_row, row_text))
                    except PlaywrightError:
                        continue
                opened = False
                for candidate, row_text in matching_rows:
                    if candidate.is_visible():
                        try:
                            log(f"{form_name}: 已匹配票据行 {row_text[:120]}")
                            candidate.dblclick(timeout=5000)
                            opened = True
                            break
                        except PlaywrightError:
                            continue
                if not opened:
                    log(f"{form_name}: 当前页未找到单据 {bill_no}，保留待处理")
                    continue
                pending.pop(bill_no)
                log(f"{form_name}: 已打开单据 {bill_no} 详情")
                try:
                    page.wait_for_timeout(700)
                    bill_date = voucher_value(row, "单据日期") or start
                    pdf_path = pdf_dir / f"{safe_filename(form_name)}-{safe_filename(bill_date)}-{safe_filename(bill_no)}.pdf"
                    expected_print_values = [bill_no, short_bill_no, bill_date]
                    if save_printbill_pdf(page, pdf_path, log, expected_print_values):
                        if generated_pdfs is not None:
                            generated_pdfs.add(pdf_path.resolve())
                        log(f"已保存 PDF: {pdf_path.name}")
                    else:
                        log(f"未生成票据 PDF: {pdf_path.name}")
                except (PlaywrightError, PlaywrightTimeoutError) as exc:
                    log(f"{form_name} 单据 PDF 失败 ({bill_no}): {exc}")
                finally:
                    close = page.get_by_role("button", name="关闭").last
                    try:
                        if close.count() and close.is_visible():
                            close.click(timeout=3000)
                            page.wait_for_timeout(300)
                    except PlaywrightError:
                        log(f"{form_name} 详情关闭失败: {bill_no}")
            if voucher_total_pages:
                selected_items = page.locator("li.paging-item.paging-item-selected:visible")
                selected_page = clean(selected_items.last.inner_text()) if selected_items.count() else ""
                if selected_page and selected_page.isdigit() and int(selected_page) >= voucher_total_pages:
                    break
            next_button = page.locator('button[paging-btn="nextPage"]:visible').last
            if not next_button.count() or next_button.is_disabled() or "disabled" in (next_button.get_attribute("class") or ""):
                break
            click_voucher_page_and_wait(page, next_button, form_name)
        for bill_no in pending:
            log(f"{form_name} 未找到可双击的票据行: {bill_no}")
    width = len(headers)
    rows_out = [row[:width] + [""] * max(0, width - len(row)) for row in rows_out]
    return headers, rows_out


def run(start: str, end: str, output: Path, selected: list[str], log, pdf_dir: Path | None = None, selected_forms: list[str] | None = None, save_excel: bool = True) -> tuple[int, int, bool]:
    if save_excel:
        output.parent.mkdir(parents=True, exist_ok=True)
    if pdf_dir is not None:
        pdf_dir.mkdir(parents=True, exist_ok=True)
    with sync_playwright() as p:
        if not debug_port_available():
            start_debug_edge()
            raise LoginRequired("已启动专用 Edge。请在新窗口登录 ERP，登录完成后再次点击“开始提取”。")
        try:
            browser = p.chromium.connect_over_cdp(f"http://127.0.0.1:{DEBUG_PORT}")
        except PlaywrightError as exc:
            raise RuntimeError(f"无法连接专用 Edge: {exc}") from exc
        pages = [page for context in browser.contexts for page in context.pages]
        if not pages:
            raise RuntimeError("未找到已连接的 Edge 页面")
        page = live_erp_page(browser)
        rows: list[SaleRow] = []
        form_sheets: dict[str, tuple[list[str], list[list[str]]]] = {}
        generated_pdfs: set[Path] = set()
        for sale_type in selected:
            page = live_erp_page(browser)
            log(f"正在读取: {sale_type}")
            module_rows = scrape_module(page, sale_type, start, end, log, pdf_dir, generated_pdfs)
            log(f"{sale_type}: 查询到 {len(module_rows)} 张销售单")
            rows.extend(module_rows)
        if not rows and not selected_forms:
            log("所选日期范围内未查询到可导出的销售单")
        if selected_forms:
            form_errors = []
            for form_name in selected_forms:
                page = live_erp_page(browser)
                log(f"正在读取: {form_name}")
                try:
                    form_sheets[form_name] = scrape_default_form(page, form_name, start, end, log, pdf_dir, generated_pdfs)
                    log(f"{form_name}: 读取 {len(form_sheets[form_name][1])} 条")
                except Exception as exc:
                    form_errors.append(f"{form_name}: {exc}")
                    log(f"{form_name} 读取失败: {exc}")
            if not form_sheets and not rows:
                raise RuntimeError("凭证类型均未读取成功: " + "；".join(form_errors))
            for error in form_errors:
                log(f"跳过失败类型: {error}")
        if save_excel and (rows or form_sheets):
            log(f"正在生成 Excel: {output}")
            export_report(rows, form_sheets, output)
            log(f"Excel 已生成: {output}")
        excel_generated = bool(save_excel and (rows or form_sheets) and output.exists())
        form_row_count = sum(len(rows_data) for _, rows_data in form_sheets.values()) if selected_forms else 0
        total_count = len(rows) + form_row_count
        pdf_count = len(generated_pdfs)
        log(f"完成，共读取 {total_count} 条（销售 {len(rows)} 条，凭证 {form_row_count} 条）；实际生成 {pdf_count} 个 PDF")
        return total_count, pdf_count, excel_generated


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.run_lock = threading.Lock()
        self.log_queue: queue.Queue[str] = queue.Queue()
        self.running = False
        self.start_button = None
        self.title("报废汽车财务数据自动化处理")
        self.geometry("860x560")
        self.minsize(760, 500)
        today = date.today().isoformat()
        self.start = tk.StringVar(value=today)
        self.end = tk.StringVar(value=today)
        app_dir = application_dir()
        self.app_dir = app_dir
        # Generated business files live below the application folder by default.
        # User-selected paths are preserved by load_settings().
        output_root = app_dir / "output"
        self.default_output = output_root / "Excel" / "销售数据.xlsx"
        self.default_pdf_dir = output_root / "PDF"
        self.output = tk.StringVar(value=str(self.default_output))
        self.pdf_dir = tk.StringVar(value=str(self.default_pdf_dir))
        self.save_pdfs = tk.BooleanVar(value=True)
        self.save_excel = tk.BooleanVar(value=True)
        self.vars = {name: tk.BooleanVar(value=True) for name in MODULES}
        self.form_vars = {name: tk.BooleanVar(value=False) for name in FORM_MODULES}
        self.settings_path = app_dir / "facas-settings.json"
        self.log_dir = app_dir / "log"
        self.run_log_path = self.log_dir / f"{date.today().isoformat()}.log"
        self.load_settings()
        self.log_box = None
        self.build()
        self._log(f"程序已启动，日志文件: {self.run_log_path}")
        self.after(100, self.flush_log_queue)
        self.after_idle(self.center_window)

    def build(self):
        frame = ttk.Frame(self, padding=(20, 18, 20, 16))
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(3, weight=1)

        heading = ttk.Label(frame, text="销售数据自动化处理", font=("Microsoft YaHei UI", 16, "bold"))
        heading.grid(row=0, column=0, sticky="w", pady=(0, 12))

        query = ttk.LabelFrame(frame, text="查询条件", padding=(14, 10))
        query.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        query.columnconfigure(1, weight=1)
        query.columnconfigure(3, weight=1)
        ttk.Label(query, text="开始日期").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=4)
        self.date_control(query, self.start, 1)
        ttk.Label(query, text="结束日期").grid(row=0, column=2, sticky="w", padx=(20, 8), pady=4)
        self.date_control(query, self.end, 3)
        selection = ttk.LabelFrame(query, text="数据范围", padding=(10, 8))
        selection.grid(row=1, column=0, columnspan=4, sticky="ew", pady=(12, 0))
        selection.columnconfigure(0, weight=1)
        selection.columnconfigure(1, weight=1)
        selection_actions = ttk.Frame(selection)
        selection_actions.grid(row=0, column=0, columnspan=2, sticky="e", pady=(0, 6))
        ttk.Button(selection_actions, text="全选全部", command=lambda: self.set_categories(True), width=9).pack(side="left", padx=(0, 6))
        ttk.Button(selection_actions, text="清空全部", command=lambda: self.set_categories(False), width=9).pack(side="left")
        sales_box = ttk.LabelFrame(selection, text="销售管理", padding=(10, 6))
        sales_box.grid(row=1, column=0, sticky="nsew", padx=(0, 6))
        vouchers_box = ttk.LabelFrame(selection, text="凭证探测器", padding=(10, 6))
        vouchers_box.grid(row=1, column=1, sticky="nsew", padx=(6, 0))
        for row, name in enumerate(MODULES):
            ttk.Checkbutton(sales_box, text=name, variable=self.vars[name]).grid(row=row, column=0, sticky="w", pady=2)
        for row, name in enumerate(FORM_MODULES):
            ttk.Checkbutton(vouchers_box, text=name, variable=self.form_vars[name]).grid(row=row, column=0, sticky="w", pady=2)

        output = ttk.LabelFrame(frame, text="输出设置", padding=(14, 10))
        output.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        output.columnconfigure(2, weight=1)
        ttk.Checkbutton(output, text="Excel", variable=self.save_excel).grid(
            row=0, column=0, sticky="w", padx=(0, 12), pady=5
        )
        ttk.Label(output, text="文件").grid(row=0, column=1, sticky="w", padx=(0, 10), pady=5)
        ttk.Entry(output, textvariable=self.output).grid(row=0, column=2, sticky="ew", pady=5)
        ttk.Button(output, text="选择文件", command=self.choose).grid(row=0, column=3, padx=(10, 0), pady=5)
        ttk.Checkbutton(output, text="PDF", variable=self.save_pdfs).grid(
            row=1, column=0, sticky="w", padx=(0, 12), pady=5
        )
        ttk.Label(output, text="目录").grid(row=1, column=1, sticky="w", padx=(0, 10), pady=5)
        ttk.Entry(output, textvariable=self.pdf_dir).grid(row=1, column=2, sticky="ew", pady=5)
        ttk.Button(output, text="选择目录", command=self.choose_pdf_dir).grid(row=1, column=3, padx=(10, 0), pady=5)

        log_frame = ttk.LabelFrame(frame, text="运行日志", padding=(8, 8))
        log_frame.grid(row=3, column=0, sticky="nsew", pady=(0, 10))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        self.log_box = tk.Listbox(
            log_frame,
            height=5,
            bg="#ffffff",
            fg="#111827",
            selectbackground="#dbeafe",
            selectforeground="#111827",
            font=("Microsoft YaHei UI", 10),
            activestyle="none",
            borderwidth=0,
            highlightthickness=0,
        )
        self.log_box.grid(row=0, column=0, sticky="nsew")
        log_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_box.yview)
        log_scroll.grid(row=0, column=1, sticky="ns")
        self.log_box.configure(yscrollcommand=log_scroll.set)

        actions = ttk.Frame(frame)
        actions.grid(row=4, column=0, sticky="ew")
        self.start_button = ttk.Button(actions, text="开始提取", command=self.start_run)
        self.start_button.pack(side="right", ipadx=16, ipady=3)

    def center_window(self):
        self.update_idletasks()
        width, height = self.winfo_width(), self.winfo_height()
        x = max(0, (self.winfo_screenwidth() - width) // 2)
        y = max(0, (self.winfo_screenheight() - height) // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")

    def date_control(self, parent, variable, column):
        box = ttk.Frame(parent)
        box.grid(row=0, column=column, padx=8, sticky="w")
        ttk.Entry(box, textvariable=variable, width=12, state="readonly").pack(side="left")
        ttk.Button(box, text="选择日期", command=lambda: self.pick_date(variable)).pack(side="left", padx=(4, 0))

    def pick_date(self, variable):
        try:
            selected = datetime.strptime(variable.get(), "%Y-%m-%d").date()
        except ValueError:
            selected = date.today()
        dialog = tk.Toplevel(self)
        dialog.title("选择日期")
        dialog.transient(self)
        dialog.grab_set()
        year = tk.IntVar(value=selected.year)
        month = tk.IntVar(value=selected.month)
        ttk.Label(dialog, text="年").grid(row=0, column=0, padx=4, pady=8)
        ttk.Spinbox(dialog, from_=2000, to=2100, textvariable=year, width=6).grid(row=0, column=1)
        ttk.Label(dialog, text="月").grid(row=0, column=2, padx=4)
        ttk.Spinbox(dialog, from_=1, to=12, textvariable=month, width=4, command=lambda: render()).grid(row=0, column=3)
        calendar_frame = ttk.Frame(dialog)
        calendar_frame.grid(row=1, column=0, columnspan=4, padx=8, pady=4)

        def render():
            for child in calendar_frame.winfo_children():
                child.destroy()
            try:
                current_year, current_month = int(year.get()), int(month.get())
                month_days = calendar.monthcalendar(current_year, current_month)
            except (ValueError, calendar.IllegalMonthError):
                return
            for col, name in enumerate(("一", "二", "三", "四", "五", "六", "日")):
                ttk.Label(calendar_frame, text=name, width=4, anchor="center").grid(row=0, column=col)
            for row_index, week in enumerate(month_days, 1):
                for col, day in enumerate(week):
                    if day:
                        ttk.Button(calendar_frame, text=str(day), width=4,
                                   command=lambda d=day: choose(d)).grid(row=row_index, column=col, padx=1, pady=1)

        def choose(day):
            variable.set(f"{int(year.get()):04d}-{int(month.get()):02d}-{day:02d}")
            dialog.destroy()

        year.trace_add("write", lambda *_: render())
        month.trace_add("write", lambda *_: render())
        render()

        # Tk places a newly-created Toplevel at the screen origin unless a
        # geometry is supplied.  Size it after rendering the calendar, then
        # center it on the main window (with a screen-centered fallback).
        dialog.update_idletasks()
        width = dialog.winfo_reqwidth()
        height = dialog.winfo_reqheight()
        parent_x = self.winfo_rootx()
        parent_y = self.winfo_rooty()
        parent_width = self.winfo_width()
        parent_height = self.winfo_height()
        if parent_width <= 1 or parent_height <= 1:
            parent_x = (dialog.winfo_screenwidth() - width) // 2
            parent_y = (dialog.winfo_screenheight() - height) // 2
        else:
            parent_x += max(0, (parent_width - width) // 2)
            parent_y += max(0, (parent_height - height) // 2)
        # Keep the dialog visible when the main window is near a screen edge.
        screen_width = dialog.winfo_screenwidth()
        screen_height = dialog.winfo_screenheight()
        parent_x = max(0, min(parent_x, screen_width - width))
        parent_y = max(0, min(parent_y, screen_height - height))
        dialog.geometry(f"{width}x{height}+{parent_x}+{parent_y}")
        dialog.focus_set()

    def choose(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if path: self.output.set(path)

    def choose_pdf_dir(self):
        path = filedialog.askdirectory()
        if path:
            self.pdf_dir.set(path)

    def set_categories(self, enabled):
        for variable in self.vars.values():
            variable.set(enabled)
        for variable in self.form_vars.values():
            variable.set(enabled)

    def load_settings(self):
        try:
            settings = json.loads(self.settings_path.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            return
        for key, variable in (("start", self.start), ("end", self.end)):
            if settings.get(key):
                variable.set(settings[key])
        # Migrate paths written by older versions, but never overwrite a
        # custom path selected by the user.
        old_default_output = self.app_dir / "销售数据.xlsx"
        old_default_pdf_dir = self.app_dir / "销售单据PDF"
        stored_output = settings.get("output")
        if stored_output:
            stored_path = Path(stored_output)
            self.output.set(str(self.default_output) if stored_path == old_default_output else stored_output)
        stored_pdf_dir = settings.get("pdf_dir")
        if stored_pdf_dir:
            stored_path = Path(stored_pdf_dir)
            self.pdf_dir.set(str(self.default_pdf_dir) if stored_path == old_default_pdf_dir else stored_pdf_dir)
        if "save_pdfs" in settings:
            self.save_pdfs.set(bool(settings["save_pdfs"]))
        if "save_excel" in settings:
            self.save_excel.set(bool(settings["save_excel"]))
        for name, variable in self.vars.items():
            if name in settings.get("categories", {}):
                variable.set(bool(settings["categories"][name]))
        for name, variable in self.form_vars.items():
            if name in settings.get("forms", {}):
                variable.set(bool(settings["forms"][name]))

    def save_settings(self):
        settings = {
            "start": self.start.get(), "end": self.end.get(),
            "output": self.output.get(), "pdf_dir": self.pdf_dir.get(),
            "save_pdfs": self.save_pdfs.get(),
            "save_excel": self.save_excel.get(),
            "categories": {name: variable.get() for name, variable in self.vars.items()},
            "forms": {name: variable.get() for name, variable in self.form_vars.items()},
        }
        try:
            self.settings_path.write_text(json.dumps(settings, ensure_ascii=False, indent=2), encoding="utf-8")
        except OSError:
            pass

    def log(self, text):
        timestamped = f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {text}"
        try:
            self.log_dir.mkdir(parents=True, exist_ok=True)
            self.run_log_path = self.log_dir / f"{date.today().isoformat()}.log"
            with self.run_log_path.open("a", encoding="utf-8") as handle:
                handle.write(timestamped + "\n")
        except OSError:
            pass
        self.log_queue.put(timestamped)

    def _log(self, text):
        self.log_box.insert("end", text.rstrip("\n") + "\n")
        self.log_box.see("end")

    def flush_log_queue(self):
        try:
            while True:
                self._log(self.log_queue.get_nowait())
        except queue.Empty:
            pass
        self.after(100, self.flush_log_queue)

    def start_run(self):
        if not self.run_lock.acquire(blocking=False):
            messagebox.showwarning("提示", "已有任务正在运行，请等待当前任务完成")
            return
        selected = [name for name, var in self.vars.items() if var.get()]
        selected_forms = [name for name, var in self.form_vars.items() if var.get()]
        if not selected and not selected_forms:
            self.run_lock.release()
            messagebox.showwarning("提示", "至少选择一个销售分类或凭证类型")
            return
        self.save_settings()
        self.running = True
        if self.start_button is not None:
            self.start_button.configure(state="disabled")
        start = self.start.get()
        end = self.end.get()
        output = Path(self.output.get())
        pdf_dir = Path(self.pdf_dir.get()) if self.save_pdfs.get() else None
        save_excel = self.save_excel.get()
        self.log(f"任务已启动，日期范围: {self.start.get()} 至 {self.end.get()}")
        self.log(f"销售分类: {', '.join(selected) if selected else '未选择'}")
        self.log(f"凭证类型: {', '.join(selected_forms) if selected_forms else '未选择'}")
        self.log(f"输出设置: Excel={'开启' if save_excel else '关闭'}；PDF={'开启' if pdf_dir is not None else '关闭'}")
        if save_excel:
            self.log(f"Excel 路径: {output}")
        if pdf_dir is not None:
            self.log(f"PDF 目录: {pdf_dir}")
        threading.Thread(target=self.worker, args=(selected, selected_forms, start, end, output, pdf_dir, save_excel), daemon=True).start()

    def worker(self, selected, selected_forms, start, end, output, pdf_dir, save_excel):
        try:
            row_count, pdf_count, excel_generated = run(start, end, output, selected, self.log, pdf_dir, selected_forms, save_excel)
            outputs = ["日期范围内未查询到可导出数据" if row_count == 0 else f"共读取 {row_count} 条数据"]
            outputs.append("已生成 1 个 Excel" if excel_generated else "未生成 Excel")
            outputs.append(f"已生成 {pdf_count} 个 PDF" if pdf_dir is not None else "未生成 PDF")
            message = "；".join(outputs) + "。"
            self.log(f"完成: {message}")
            self.log(f"统计: 实际读取 {row_count} 条，Excel {'1' if excel_generated else '0'} 个，PDF {pdf_count if pdf_dir is not None else '0'} 个")
            if excel_generated:
                self.log(f"Excel 文件: {output}")
            if pdf_dir is not None:
                self.log(f"PDF 目录: {pdf_dir}")
            open_path = output.parent if excel_generated else pdf_dir
            if open_path is not None:
                self.after(0, lambda path=open_path: open_output_directory(path))
            self.after(0, lambda text=message: messagebox.showinfo("完成", text))
        except LoginRequired as exc:
            message = str(exc)
            self.log(message)
            self.after(0, lambda text=message: messagebox.showinfo("请先登录 ERP", text))
        except Exception as exc:
            error_text = str(exc)
            self.log(f"失败: {error_text}")
            self.log(f"错误类型: {type(exc).__name__}")
            self.log(traceback.format_exc().strip())
            self.after(0, lambda text=error_text: messagebox.showerror("失败", text))
        finally:
            self.run_lock.release()
            self.running = False
            self.after(0, lambda: self.start_button.configure(state="normal") if self.start_button is not None else None)


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--start"); parser.add_argument("--end"); parser.add_argument("--output", type=Path); parser.add_argument("--pdf-dir", type=Path)
    args = parser.parse_args()
    if args.start and args.end:
        cli_output = args.output or (Path.cwd() / "output" / "Excel" / "销售数据.xlsx")
        run(args.start, args.end, cli_output, list(MODULES), print, args.pdf_dir)
    else:
        App().mainloop()
