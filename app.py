from __future__ import annotations

import argparse
import base64
import calendar
import ctypes
import json
import queue
import re
import socket
import subprocess
import threading
import time
import traceback
import tkinter as tk
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from playwright.sync_api import Error as PlaywrightError, Page, TimeoutError as PlaywrightTimeoutError, sync_playwright
from pywinauto import Desktop, keyboard


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


def start_debug_edge() -> None:
    """Open a visible Edge window with a persistent automation profile."""
    edge_paths = (
        Path(r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"),
        Path(r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"),
    )
    edge = next((item for item in edge_paths if item.exists()), None)
    if edge is None:
        raise RuntimeError("未找到 Microsoft Edge，请安装 Edge 后重试")
    profile = Path.cwd() / ".facas-edge-profile"
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


def copy_to_clipboard(value: str) -> None:
    """Put Unicode text on the Windows clipboard for native save dialogs."""
    data = (value + "\0").encode("utf-16-le")
    kernel32 = ctypes.windll.kernel32
    user32 = ctypes.windll.user32
    kernel32.GlobalAlloc.argtypes = (ctypes.c_uint, ctypes.c_size_t)
    kernel32.GlobalAlloc.restype = ctypes.c_void_p
    kernel32.GlobalLock.argtypes = (ctypes.c_void_p,)
    kernel32.GlobalLock.restype = ctypes.c_void_p
    kernel32.GlobalUnlock.argtypes = (ctypes.c_void_p,)
    kernel32.GlobalFree.argtypes = (ctypes.c_void_p,)
    user32.OpenClipboard.argtypes = (ctypes.c_void_p,)
    user32.SetClipboardData.argtypes = (ctypes.c_uint, ctypes.c_void_p)
    user32.SetClipboardData.restype = ctypes.c_void_p
    global_memory = kernel32.GlobalAlloc(0x0002, len(data))
    if not global_memory:
        raise RuntimeError("无法分配 Windows 剪贴板内存")
    pointer = kernel32.GlobalLock(global_memory)
    if not pointer:
        kernel32.GlobalFree(global_memory)
        raise RuntimeError("无法锁定 Windows 剪贴板内存")
    ctypes.memmove(pointer, data, len(data))
    kernel32.GlobalUnlock(global_memory)
    if not user32.OpenClipboard(None):
        raise RuntimeError("无法打开 Windows 剪贴板")
    try:
        user32.EmptyClipboard()
        if not user32.SetClipboardData(13, global_memory):  # CF_UNICODETEXT
            raise RuntimeError("无法写入 Windows 剪贴板")
        global_memory = None  # Clipboard now owns this handle.
    finally:
        user32.CloseClipboard()
        if global_memory:
            kernel32.GlobalFree(global_memory)


def save_print_dialog_pdf(path: Path) -> None:
    """Fill the Windows 'Save Print Output As' dialog from Microsoft Print to PDF."""
    desktop = Desktop(backend="uia")
    desktop_win32 = Desktop(backend="win32")
    dialog = None
    deadline = time.monotonic() + 15
    while time.monotonic() < deadline:
        for candidate_desktop in (desktop, desktop_win32):
            for window in candidate_desktop.windows():
                if not window.is_visible():
                    continue
                try:
                    title = window.window_text().lower()
                    is_save_dialog = any(token in title for token in ("save", "另存", "保存"))
                    has_filename = window.child_window(auto_id="FileNameControl", control_type="ComboBox").exists(timeout=0.1)
                    if is_save_dialog or has_filename:
                        dialog = window
                        break
                except Exception:
                    continue
            if dialog is not None:
                break
        if dialog is not None:
            break
        time.sleep(0.2)
    if dialog is None:
        # The ERP's LODOP dialog can deny UI Automation access. Its filename
        # field receives initial focus, so use clipboard paste on the active
        # native dialog without inspecting its controls.
        copy_to_clipboard(str(path))
        keyboard.send_keys("^a")
        keyboard.send_keys("^v")
        keyboard.send_keys("{ENTER}")
        return

    dialog.set_focus()
    try:
        filename = dialog.child_window(auto_id="FileNameControl", control_type="ComboBox")
        edit = filename.child_window(control_type="Edit")
        if edit.exists(timeout=1):
            edit.set_edit_text(str(path))
        else:
            filename.set_edit_text(str(path))
        save_button = dialog.child_window(title_re=".*(保存|Save).*", control_type="Button")
        save_button.click()
    except Exception:
        # Traditional Win32 print dialogs may not expose UIA identifiers.
        keyboard.send_keys("^a")
        keyboard.send_keys(str(path), with_spaces=True)
        keyboard.send_keys("{ENTER}")
    try:
        confirm = desktop.window(title_re=".*(确认另存为|Confirm Save As|确认保存).*", visible_only=True)
        confirm.wait("visible", timeout=2)
        confirm.child_window(title_re="^(是|Yes)$", control_type="Button").click()
    except Exception:
        pass


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
        if best_markup is not None and (not expected or best_score > 0):
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


def save_printbill_pdf(page: Page, path: Path, log) -> bool:
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
        expected_values = print_payload_values(payload)
        # The response often contains only the fixed template. Prefer the
        # fully populated #bill1 rendered by the ERP print iframe.
        markup = page_print_markup(page, expected_values)
        if not markup:
            response_markup = find_print_markup(payload)
            if response_markup and any(value in response_markup for value in expected_values):
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


def detail_summary(page: Page) -> tuple[str, str, str, str]:
    """Read only the first detail row in the open bill."""
    # Detail tables are identified by their header, not by global table order.
    candidates = page.locator("table.yc-view-grid-table")
    table = None
    for i in range(candidates.count()):
        candidate = candidates.nth(i)
        if (candidate.is_visible()
                and candidate.locator("tr").first.locator("td").count() >= 15):
            table = candidate
            break
    if table is None:
        return "", "", "", ""
    table.wait_for(state="visible", timeout=10000)
    rows = table.locator("tr")
    if not rows.count():
        rows = table.locator("tr").nth(1)
    if not rows.count():
        return "", "", "", ""
    cells = rows.first.locator("td")
    values = [clean(cells.nth(i).inner_text()) for i in range(cells.count())]
    amount_total = 0.0
    tax_total = 0.0
    for row in rows.all():
        cells = row.locator("td")
        vals = [clean(cells.nth(i).inner_text()) for i in range(cells.count())]
        try:
            amount_total += float(vals[11].replace(",", ""))
        except (ValueError, IndexError):
            pass
        try:
            tax_total += float(vals[12].replace(",", ""))
        except (ValueError, IndexError):
            pass
    return (
        values[1] if len(values) > 1 else "",
        values[2] if len(values) > 2 else "",
        f"{amount_total:,.2f}" if amount_total else "",
        f"{tax_total:,.2f}" if tax_total else "",
    )


def data_grid(page: Page):
    """The ERP renders a header grid, a data grid, and a subtotal grid."""
    grids = page.locator("table.yc-view-grid-table")
    for i in range(grids.count()):
        grid = grids.nth(i)
        if (grid.is_visible() and grid.locator("tr").count() >= 1
                and grid.locator("tr").first.locator("td").count() >= 8):
            return grid
    for i in range(grids.count()):
        if grids.nth(i).is_visible():
            return grids.nth(i)
    return grids.last


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


def apply_date_filters(filter_table, start: str, end: str) -> bool:
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
        search.click()
        return True
    return False


def apply_sale_date_filters(filter_table, start: str, end: str) -> bool:
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
        search.click()
        return True
    return False
    # Identify the two sales-date controls by their actual labels. Do not use
    # generic "日期" matching: 回收单号 may contain a date value as well.
    date_controls = []
    containers = filter_table.locator("td:visible")
    if not containers.count():
        containers = filter_table.locator("tr:visible")
    label_specs = (("销售日期从", start), ("销售日期到", end))
    for label, value in label_specs:
        found = False
        for index in range(containers.count()):
            container = containers.nth(index)
            text = clean(container.inner_text())
            if label not in text:
                continue
            row_checkbox = container.locator("input[type='checkbox']:visible")
            row_input = container.locator("input:not([type='checkbox']):visible")
            if row_checkbox.count() and row_input.count():
                date_controls.append((row_checkbox.first, row_input.first, value))
                found = True
                break
        if not found:
            return False
    # Clear every existing condition, then enable only the two date rows.
    for index in range(checkboxes.count()):
        checkbox = checkboxes.nth(index)
        if checkbox.is_checked():
            checkbox.uncheck()
    for checkbox, date_input, value in date_controls:
        if not checkbox.is_checked():
            checkbox.check()
        date_input.fill(value)
        date_input.press("Enter")
    # Some ERP controls re-render after date input and restore their previous
    # state. Re-check every visible checkbox and leave only the two date ones.
    date_boxes = []
    for checkbox, _, _ in date_controls:
        box = checkbox.bounding_box()
        if box:
            date_boxes.append((round(box["x"]), round(box["y"])))
    for index in range(checkboxes.count()):
        checkbox = checkboxes.nth(index)
        box = checkbox.bounding_box()
        position = (round(box["x"]), round(box["y"])) if box else None
        if position not in date_boxes and checkbox.is_checked():
            checkbox.uncheck()
    search = filter_table.get_by_role("button", name="搜索")
    if search.count():
        search.click()
        return True
    return False


def find_voucher_rows(value):
    """Find vc_bill_list rows inside an ERP response payload."""
    if isinstance(value, dict):
        if isinstance(value.get("rowsData"), list) and value.get("name") == "vc_bill_list":
            return value["rowsData"]
        for child in value.values():
            found = find_voucher_rows(child)
            if found is not None:
                return found
    elif isinstance(value, list):
        for child in value:
            found = find_voucher_rows(child)
            if found is not None:
                return found
    return None


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
    # The menu contains both a navigation item and a tab. Clicking the first
    # exact text item follows the same route as a normal user click.
    if page.is_closed():
        raise RuntimeError("ERP 页面已关闭，请重新登录专用 Edge")
    expand_menu_section(page, "销售管理", sale_type)
    click_visible_menu_item(page, sale_type)
    page.wait_for_timeout(500)
    table = data_grid(page)
    table.wait_for(state="visible", timeout=15000)

    # Prototype date filtering: fill the two date controls and enable their
    # adjacent checkboxes. The application then performs the normal search.
    filter_tables = page.locator("table.yc-view-free-table:visible")
    filter_table = filter_tables.first if filter_tables.count() else page.locator("table.yc-view-free-table").first
    if apply_sale_date_filters(filter_table, start, end):
        page.wait_for_timeout(1000)

    output: list[SaleRow] = []
    listed_count = 0
    skipped_payment_count = 0
    page_no = 1
    seen_sale_nos: set[str] = set()
    while True:
        if page.is_closed():
            available = [candidate for candidate in page.context.pages if not candidate.is_closed()]
            if not available:
                raise RuntimeError("打印后浏览器页面已关闭，无法继续读取销售列表")
            page = available[0]
            page.wait_for_timeout(500)
        grid = data_grid(page)
        rows = grid.locator("tr")
        log(f"{sale_type}: 正在读取第 {page_no} 页")
        for i in range(rows.count()):
            cells = rows.nth(i).locator("td")
            vals = [clean(cells.nth(j).inner_text()) for j in range(cells.count())]
            if len(vals) < 8 or not vals[1] or not re.search(r"\d{4}-\d{2}-\d{2}", vals[2]):
                continue
            sale_no, sale_date = vals[1], vals[2]
            if sale_no in seen_sale_nos:
                continue
            seen_sale_nos.add(sale_no)
            listed_count += 1
            payment_date = vals[3] if len(vals) > 3 else ""
            customer = vals[4] if len(vals) > 4 else ""
            if not payment_date or not re.search(r"\d{4}-\d{2}-\d{2}", payment_date):
                skipped_payment_count += 1
                continue
            amount = vals[5] if len(vals) > 5 else ""
            weight = vals[6] if len(vals) > 6 else ""
            material = recovery_no = remark = tax = ""
            close = None
            try:
                rows.nth(i).get_by_text("查看", exact=True).click()
                close = page.get_by_role("button", name="关闭").last
                close.wait_for(state="visible", timeout=10000)
                page.wait_for_timeout(700)
                material, recovery_no, calculated_amount, tax = detail_summary(page)
                if not amount:
                    amount = calculated_amount
                remark_box = page.locator("textarea:visible").last
                remark = clean(remark_box.input_value()) if remark_box.count() else ""
                if pdf_dir is not None:
                    pdf_dir.mkdir(parents=True, exist_ok=True)
                    pdf_path = pdf_dir / f"{safe_filename(sale_type)}-{safe_filename(sale_date)}-{safe_filename(sale_no)}.pdf"
                    # Capture the ERP ticket-print response, rather than
                    # printing the underlying detail form.
                    if not save_printbill_pdf(page, pdf_path, log):
                        log(f"未生成票据 PDF: {pdf_path.name}")
                    else:
                        if generated_pdfs is not None:
                            generated_pdfs.add(pdf_path.resolve())
                        log(f"已保存 PDF: {pdf_path.name}")
            except (PlaywrightTimeoutError, PlaywrightError, IndexError, OSError):
                log(f"详情读取失败: {sale_no}")
            finally:
                try:
                    if close is not None and not page.is_closed() and close.count() and close.is_visible():
                        close.click(timeout=3000)
                        page.wait_for_timeout(300)
                except PlaywrightError:
                    log(f"详情关闭失败: {sale_no}")
            output.append(SaleRow(sale_type, sale_no, sale_date, payment_date, material, recovery_no, amount, tax, weight, remark))

        next_button = page.locator('button[paging-btn="nextPage"]')
        if not next_button.count() or next_button.is_disabled() or "disabled" in (next_button.get_attribute("class") or ""):
            break
        first_no = ""
        if rows.count() and rows.first.locator("td").count() > 1:
            first_no = clean(rows.first.locator("td").nth(1).inner_text())
        next_button.click()
        try:
            page.wait_for_timeout(800)
            for _ in range(20):
                new_grid = data_grid(page)
                new_rows = new_grid.locator("tr")
                if new_rows.count() and new_rows.first.locator("td").count() > 1:
                    new_no = clean(new_rows.first.locator("td").nth(1).inner_text())
                    if new_no and new_no != first_no:
                        break
                page.wait_for_timeout(300)
            else:
                break
        except PlaywrightTimeoutError:
            break
        page_no += 1
    log(f"{sale_type}: 列表销售单 {listed_count} 张，因收款日期为空跳过 {skipped_payment_count} 张，导出 {len(output)} 张")
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


def export_xlsx(rows: list[SaleRow], path: Path) -> None:
    export_report(rows, {}, path)


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
    # The last path step may only expose the form label; click it explicitly.
    click_visible_menu_item(page, form_name)
    page.wait_for_timeout(700)
    table = data_grid(page)
    table.wait_for(state="visible", timeout=15000)
    response_rows = []

    def capture_response(response):
        if "erp.bfcgj.com" not in response.url:
            return
        try:
            payload = response.json()
        except Exception:
            return
        rows = find_voucher_rows(payload)
        if rows is not None:
            response_rows.extend(rows)

    page.on("response", capture_response)
    filter_tables = page.locator("table.yc-view-free-table:visible")
    filter_table = filter_tables.first if filter_tables.count() else page.locator("table.yc-view-free-table").first
    if apply_date_filters(filter_table, start, end):
        page.wait_for_timeout(900)
    else:
        search = filter_table.get_by_role("button", name="搜索")
        if search.count():
            search.click()
            page.wait_for_timeout(900)

    # Confirmed visible order, mapped to the vc_bill_list response fields.
    headers: list[str] = [
        "单据日期", "完成日期", "票据号码", "对方单位",
        "金额", "结算方式", "制单",
    ]
    rows_out: list[list[str]] = []
    seen: set[tuple[str, ...]] = set()
    page_no = 1
    while True:
        table = data_grid(page)
        rows = table.locator("tbody tr")
        if not rows.count():
            rows = table.locator("tr")
        log(f"{form_name}: 正在读取第 {page_no} 页")
        for i in range(rows.count()):
            cells = rows.nth(i).locator("td")
            values = [clean(cells.nth(j).inner_text()) for j in range(cells.count())]
            if not values or not any(values):
                continue
            key = tuple(values)
            if key not in seen:
                seen.add(key)
                rows_out.append(values)
        next_button = page.locator('button[paging-btn="nextPage"]')
        if not next_button.count() or next_button.is_disabled() or "disabled" in (next_button.get_attribute("class") or ""):
            break
        before = tuple(rows_out[-1]) if rows_out else ()
        next_button.click()
        page.wait_for_timeout(800)
        if rows_out and tuple(clean(x) for x in data_grid(page).locator("tr").last.locator("td").all_inner_texts()) == before:
            break
        page_no += 1
    if not response_rows:
        raise RuntimeError("未捕获 vc_bill_list 响应，已停止导出以避免生成错误列数据")
    rows_out = []
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
        pdf_dir.mkdir(parents=True, exist_ok=True)
        pending = {}
        for row in response_rows:
            bill_no = clean(str(row.get("c0", "")))
            if bill_no:
                pending.setdefault(bill_no, row)

        # List extraction ends on the final page. Return to the beginning and
        # process the visible bills page by page so earlier pages are not lost.
        previous_button = page.locator('button[paging-btn="prevPage"]')
        for _ in range(200):
            if not previous_button.count() or previous_button.is_disabled() or "disabled" in (previous_button.get_attribute("class") or ""):
                break
            previous_button.click()
            page.wait_for_timeout(800)

        while pending:
            grid = data_grid(page)
            for bill_no, row in list(pending.items()):
                matching_row = grid.locator("tr").filter(has_text=bill_no)
                opened = False
                for index in range(matching_row.count()):
                    candidate = matching_row.nth(index)
                    if candidate.is_visible():
                        try:
                            candidate.dblclick(timeout=5000)
                            opened = True
                            break
                        except PlaywrightError:
                            continue
                if not opened:
                    continue
                pending.pop(bill_no)
                try:
                    page.wait_for_timeout(700)
                    bill_date = voucher_value(row, "单据日期") or start
                    pdf_path = pdf_dir / f"{safe_filename(form_name)}-{safe_filename(bill_date)}-{safe_filename(bill_no)}.pdf"
                    if save_printbill_pdf(page, pdf_path, log):
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
            next_button = page.locator('button[paging-btn="nextPage"]')
            if not next_button.count() or next_button.is_disabled() or "disabled" in (next_button.get_attribute("class") or ""):
                break
            next_button.click()
            page.wait_for_timeout(800)
        for bill_no in pending:
            log(f"{form_name} 未找到可双击的票据行: {bill_no}")
    if not headers and rows_out:
        headers = [f"字段{i + 1}" for i in range(len(rows_out[0]))]
    width = len(headers)
    rows_out = [row[:width] + [""] * max(0, width - len(row)) for row in rows_out]
    return headers, rows_out


def export_form_sheets(sheets: dict[str, tuple[list[str], list[list[str]]]], path: Path) -> None:
    export_report([], sheets, path)


def run(start: str, end: str, output: Path, selected: list[str], log, pdf_dir: Path | None = None, selected_forms: list[str] | None = None, save_excel: bool = True) -> tuple[int, int]:
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
        if not rows:
            if not selected_forms:
                raise RuntimeError("未查询到销售单，未生成 Excel。请确认日期范围、销售分类和 Edge 页面登录状态。")
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
            if not form_sheets and not rows and save_excel:
                raise RuntimeError("凭证类型均未读取成功: " + "；".join(form_errors))
            for error in form_errors:
                log(f"跳过失败类型: {error}")
        if save_excel and (rows or form_sheets):
            export_report(rows, form_sheets, output)
        form_row_count = sum(len(rows_data) for _, rows_data in form_sheets.values()) if selected_forms else 0
        total_count = len(rows) + form_row_count
        pdf_count = len(generated_pdfs)
        log(f"完成，共读取 {total_count} 条（销售 {len(rows)} 条，凭证 {form_row_count} 条）；实际生成 {pdf_count} 个 PDF")
        return total_count, pdf_count


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.run_lock = threading.Lock()
        self.log_queue: queue.Queue[str] = queue.Queue()
        self.running = False
        self.start_button = None
        self.title("报废汽车财务数据自动化处理")
        self.geometry("860x680")
        self.minsize(760, 590)
        today = date.today().isoformat()
        self.start = tk.StringVar(value=today)
        self.end = tk.StringVar(value=today)
        app_dir = Path(__file__).resolve().parent
        self.app_dir = app_dir
        self.output = tk.StringVar(value=str(app_dir / "销售数据.xlsx"))
        self.pdf_dir = tk.StringVar(value=str(app_dir / "销售单据PDF"))
        self.save_pdfs = tk.BooleanVar(value=True)
        self.save_excel = tk.BooleanVar(value=True)
        self.vars = {name: tk.BooleanVar(value=True) for name in MODULES}
        self.form_vars = {name: tk.BooleanVar(value=False) for name in FORM_MODULES}
        self.settings_path = app_dir / "facas-settings.json"
        self.run_log_path = app_dir / "facas-run.log"
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
        output.columnconfigure(1, weight=1)
        switch_box = ttk.Frame(output)
        switch_box.grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 6))
        ttk.Checkbutton(switch_box, text="生成 Excel", variable=self.save_excel).pack(side="left", padx=(0, 24))
        ttk.Checkbutton(switch_box, text="保存单据 PDF", variable=self.save_pdfs).pack(side="left")
        ttk.Label(output, text="Excel 文件").grid(row=1, column=0, sticky="w", padx=(0, 10), pady=5)
        ttk.Entry(output, textvariable=self.output).grid(row=1, column=1, sticky="ew", pady=5)
        ttk.Button(output, text="选择文件", command=self.choose).grid(row=1, column=2, padx=(10, 0), pady=5)
        ttk.Label(output, text="PDF 目录").grid(row=2, column=0, sticky="w", padx=(0, 10), pady=5)
        ttk.Entry(output, textvariable=self.pdf_dir).grid(row=2, column=1, sticky="ew", pady=5)
        ttk.Button(output, text="选择目录", command=self.choose_pdf_dir).grid(row=2, column=2, padx=(10, 0), pady=5)

        log_frame = ttk.LabelFrame(frame, text="运行日志", padding=(8, 8))
        log_frame.grid(row=3, column=0, sticky="nsew", pady=(0, 10))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        self.log_box = tk.Listbox(
            log_frame,
            height=12,
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

    def set_range(self, start, end):
        self.start.set(start.isoformat())
        self.end.set(end.isoformat())

    def set_today(self):
        today = date.today(); self.set_range(today, today)

    def set_yesterday(self):
        yesterday = date.today() - timedelta(days=1); self.set_range(yesterday, yesterday)

    def set_this_month(self):
        today = date.today(); self.set_range(today.replace(day=1), today)

    def set_last_month(self):
        today = date.today().replace(day=1)
        last_day = today - timedelta(days=1)
        self.set_range(last_day.replace(day=1), last_day)

    def load_settings(self):
        try:
            settings = json.loads(self.settings_path.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            return
        for key, variable in (("start", self.start), ("end", self.end), ("output", self.output), ("pdf_dir", self.pdf_dir)):
            if settings.get(key):
                variable.set(settings[key])
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
            with self.run_log_path.open("a", encoding="utf-8") as handle:
                handle.write(timestamped + "\n")
        except OSError:
            pass
        self.log_queue.put(timestamped)

    def _log(self, text):
        self.log_box.insert("end", text)
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
        threading.Thread(target=self.worker, args=(selected, selected_forms, start, end, output, pdf_dir, save_excel), daemon=True).start()

    def worker(self, selected, selected_forms, start, end, output, pdf_dir, save_excel):
        try:
            row_count, pdf_count = run(start, end, output, selected, self.log, pdf_dir, selected_forms, save_excel)
            outputs = [f"共读取 {row_count} 条数据"]
            outputs.append("已生成 1 个 Excel" if save_excel else "未生成 Excel")
            outputs.append(f"已生成 {pdf_count} 个 PDF" if pdf_dir is not None else "未生成 PDF")
            message = "；".join(outputs) + "。"
            self.log(f"完成: {message}")
            self.log(f"统计: 实际读取 {row_count} 条，Excel {'1' if save_excel else '0'} 个，PDF {pdf_count if pdf_dir is not None else '0'} 个")
            if save_excel:
                self.log(f"Excel 文件: {output}")
            if pdf_dir is not None:
                self.log(f"PDF 目录: {pdf_dir}")
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
        run(args.start, args.end, args.output or Path("销售数据.xlsx"), list(MODULES), print, args.pdf_dir)
    else:
        App().mainloop()
