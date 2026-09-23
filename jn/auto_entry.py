# -*- coding: utf-8 -*-
"""银行流水凭证自动录入 - 独立脚本"""

from __future__ import annotations

import os
import sys
import time
from pathlib import Path
from typing import Any, Callable

import pandas as pd
from playwright.sync_api import Locator, Page, sync_playwright

CDP_URL = "http://localhost:9222"
DEFAULT_WAIT = 10000

COL_DATE = "交易日期"
COL_SUMMARY = "摘要"
COL_COUNTERPARTY = "对方账户名称"
COL_DEBIT = "借方发生额"
COL_CREDIT = "贷方发生额"
COL_DIRECTION = "资金方向"
COL_VOUCHER_NO = "凭证号"
COL_SKIP = "不需要录入凭证"

SUBJECT_PAY_DEBIT = "220201"
SUBJECT_RECEIVE_CREDIT = "112201"
DEFAULT_BANK_CODE = "10029334MZ"  # 银行科目码：会随录入的银行账户变化，由前端传入

FALLBACK_SEARCH = "不分明细"


LogFunction = Callable[[str], None]


def log(msg: str) -> None:
    print(f"[{time.strftime('%H:%M:%S')}] {msg}", flush=True)


def load_rows(excel_path: str | Path) -> list[dict[str, Any]]:
    df = pd.read_excel(excel_path)
    for col in [COL_DEBIT, COL_CREDIT]:
        if col in df.columns:
            df[col] = pd.to_numeric(
                df[col].astype(str).str.replace(",", "", regex=False),
                errors="coerce",
            ).fillna(0)
    rows = []
    for _, row in df.iterrows():
        debit = float(row[COL_DEBIT]) if COL_DEBIT in df.columns else 0
        credit = float(row[COL_CREDIT]) if COL_CREDIT in df.columns else 0
        if COL_DIRECTION in df.columns and not pd.isna(row[COL_DIRECTION]):
            direction = str(row[COL_DIRECTION]).strip()
        else:
            direction = "付" if debit > 0 else ("收" if credit > 0 else "")
        voucher_no = ""
        if COL_VOUCHER_NO in df.columns and not pd.isna(row[COL_VOUCHER_NO]):
            raw_voucher = row[COL_VOUCHER_NO]
            if isinstance(raw_voucher, (int, float)):
                voucher_no = str(int(raw_voucher))
            else:
                voucher_no = str(raw_voucher).strip()
        skip = ""
        if COL_SKIP in df.columns and not pd.isna(row[COL_SKIP]):
            raw_skip = str(row[COL_SKIP]).strip()
            try:
                # Excel 里填的数字 1 会读成 1.0，统一成 "1"
                if float(raw_skip) == 1.0:
                    skip = "1"
                else:
                    skip = raw_skip
            except Exception:
                skip = raw_skip
        rows.append(
            {
                "date": str(row[COL_DATE]).strip()[:10],
                "summary": (
                    "" if pd.isna(row[COL_SUMMARY]) else str(row[COL_SUMMARY]).strip()
                ),
                "counterparty": (
                    ""
                    if pd.isna(row[COL_COUNTERPARTY])
                    else str(row[COL_COUNTERPARTY]).strip()
                ),
                "debit": debit,
                "credit": credit,
                "direction": direction,
                "voucher_no": voucher_no,
                "skip": skip,
            }
        )
    return rows


class Locators:
    MENU_FINANCE = "#main > div > div > div.yc-view.yc-view-container.desk-container > div.yc-view-area.side-navigation-content > div > div > div:nth-child(1) > div.flex-item-header > div"
    MENU_VOUCHER = "#main > div > div > div.yc-view.yc-view-container.desk-container > div.yc-view-area.side-navigation-content > div > div > div:nth-child(1) > div.flex-item-body > div:nth-child(1) > div.flex-item-header > div"
    MENU_ENTRY = "#main > div > div > div.yc-view.yc-view-container.desk-container > div.yc-view-area.side-navigation-content > div > div > div:nth-child(1) > div.flex-item-body > div:nth-child(1) > div.flex-item-body > div:nth-child(1) > div > div"
    BTN_ADD = "#main > div > div > div.yc-view.yc-view-container.desk-container > div.yc-view-area.viewframework-body > div > div:nth-child(2) > div > div > div.yc-view.yc-view-button > div.yc-align-left > button:nth-child(2)"
    BTN_SAVE = "body > div:nth-child(4) > div.yc-view-dialog-body > div:nth-child(1) > div.yc-view-area.cols16 > div > div > button:nth-child(1)"
    DATE_INPUT = "body > div:nth-child(4) > div.yc-view-dialog-body > div:nth-child(2) > div > div:nth-child(1) > div > div.yc-view.yc-view-free.full-panel.finance-free > div.yc-view-body.yc-view-free-body > table > tbody > tr:nth-child(1) > td:nth-child(2) > div > div > input[type=text]"
    VOUCHER_NO_INPUT = "body > div:nth-child(4) > div.yc-view-dialog-body > div:nth-child(2) > div > div:nth-child(1) > div > div.yc-view.yc-view-free.full-panel.finance-free > div.yc-view-body.yc-view-free-body > table > tbody > tr:nth-child(2) > td:nth-child(4) > div > div > input[type=text]"
    GRID = "#\\31 004 > div > div.yc-view-grid-body > table > tbody"
    BTN_PLUS = "body > div:nth-child(4) > div.yc-view-dialog-body > div:nth-child(2) > div > div:nth-child(2) > div:nth-child(2) > div > div.yc-view-body.yc-view-tabs-body > div > div:nth-child(2) > div > div > button:nth-child(1)"
    MAGNIFIER = "body > div:nth-child(10) > div.yc-view-dialog-body > div > div.yc-view-area.cols16 > div.yc-view.yc-view-free > div > table > tbody > tr > td:nth-child(2) > div > div > div.field-btn-right > button"
    # 凭证录入界面的"关闭"按钮（标题栏右边的 ×）
    BTN_CLOSE_VOUCHER = "body > div:nth-child(4) > div.yc-view-dialog-header > div > i"


def click(
    page: Page, selector: str, desc: str = "", wait: int = 1500, force: bool = False
) -> None:
    log(f"  点击：{desc or selector}")
    el = page.locator(selector).first
    el.wait_for(state="visible", timeout=DEFAULT_WAIT)
    if force:
        el.click(force=True)
    else:
        el.click()
    page.wait_for_timeout(wait)


def fill(page: Page, selector: str, text: str, desc: str = "", wait: int = 500) -> None:
    log(f"  填入 [{text}] → {desc}")
    el = page.locator(selector).first
    el.wait_for(state="visible", timeout=DEFAULT_WAIT)
    el.click()
    el.fill("")
    el.type(str(text), delay=30)
    page.wait_for_timeout(wait)


def fill_and_confirm_subject(page: Page, row_index: int, subject_code: str) -> None:
    log(f"  填写第 {row_index} 行会计科目：{subject_code}")
    selector = f"{Locators.GRID} > tr:nth-child({row_index}) > td:nth-child(3) > div > div > div.field-input > textarea"
    el = page.locator(selector).first
    el.wait_for(state="visible", timeout=DEFAULT_WAIT)
    el.click()
    el.fill("")
    el.type(subject_code, delay=50)
    page.wait_for_timeout(1200)
    page.keyboard.press("Enter")
    page.wait_for_timeout(500)


def fill_summary(page: Page, row_index: int, text: str) -> None:
    log(f"  填写第 {row_index} 行摘要：{text[:30]}{'...' if len(text) > 30 else ''}")
    selector = f"{Locators.GRID} > tr:nth-child({row_index}) > td:nth-child(2)"
    el = page.locator(selector).first
    el.wait_for(state="visible", timeout=DEFAULT_WAIT)
    el.click()
    page.wait_for_timeout(300)
    page.keyboard.type(text, delay=20)
    page.keyboard.press("Tab")
    page.wait_for_timeout(400)


def fill_amount(
    page: Page, row_index: int, col_index: int, amount: float, press_tab: bool = True
) -> None:
    log(f"  填写第 {row_index} 行金额：{amount:.2f}")
    selector = (
        f"{Locators.GRID} > tr:nth-child({row_index}) > td:nth-child({col_index})"
    )
    el = page.locator(selector).first
    el.wait_for(state="visible", timeout=DEFAULT_WAIT)
    el.click()
    page.wait_for_timeout(200)
    page.keyboard.type(f"{amount:.2f}", delay=20)
    if press_tab:
        page.keyboard.press("Tab")
    page.wait_for_timeout(400)


# ========== 弹窗相关 ==========
def find_dialog(page: Page, timeout: int = 8000) -> Locator | None:
    end = time.time() + timeout / 1000
    while time.time() < end:
        dialogs = page.locator("div.yc-view-dialog")
        n = dialogs.count()
        for i in range(n - 1, -1, -1):
            try:
                dlg = dialogs.nth(i)
                if dlg.is_visible():
                    return dlg
            except Exception:
                continue
        page.wait_for_timeout(300)
    return None


def is_project_dialog_open(page: Page) -> bool:
    """仅判断“核算项目列表”弹窗是否打开。"""
    try:
        title = page.locator("text=核算项目列表").first
        return title.count() > 0 and title.is_visible()
    except Exception:
        return False


def close_project_dialog_if_open(page: Page) -> bool:
    """仅在“核算项目列表”弹窗打开时关闭它，不操作主界面。"""
    if not is_project_dialog_open(page):
        return False
    log("  检测到核算项目弹窗未关，尝试关闭")
    try:
        dialogs = page.locator("div.yc-view-dialog")
        for i in range(dialogs.count()):
            dlg = dialogs.nth(i)
            try:
                if not dlg.is_visible():
                    continue
                title = dlg.locator("text=核算项目列表")
                if title.count() == 0:
                    continue
                cancel = dlg.locator("button").filter(has_text="取消").first
                if cancel.count() > 0 and cancel.is_visible():
                    cancel.click(timeout=2000)
                else:
                    dlg.locator("i.fa-close").first.click(timeout=2000)
                page.wait_for_timeout(600)
                log("  已关闭核算项目弹窗")
                return True
            except Exception:
                continue
    except Exception as e:
        log(f"  关闭核算项目弹窗失败：{str(e)[:80]}")
    return False


def close_voucher_window(page: Page) -> None:
    """关闭当前凭证窗口，为录入下一条凭证做准备。"""
    log("  关闭当前凭证录入窗口")
    try:
        # 先判断是否还开着（有些情况保存后自动关了）
        btn = page.locator(Locators.BTN_CLOSE_VOUCHER).first
        if btn.count() == 0:
            log("  凭证录入窗口已自动关闭")
            return
        try:
            btn.click(timeout=3000)
            page.wait_for_timeout(1000)
            log("  凭证录入窗口已关闭")
        except Exception:
            # 用 JS 强点
            btn.evaluate("el => el.click()")
            page.wait_for_timeout(1000)
            log("  凭证录入窗口已关闭（JS 点击）")
    except Exception as e:
        log(f"  关闭凭证录入窗口失败：{str(e)[:80]}")


def select_counterparty(
    page: Page, counterparty_name: str, fallback: str = FALLBACK_SEARCH
) -> None:
    log(f"  打开核算项目弹窗，搜索：{counterparty_name or '(空)'}")
    click(page, Locators.BTN_PLUS, "核算项目加号", wait=2500)

    dialog = find_dialog(page, timeout=8000)
    if dialog is None:
        raise RuntimeError("找不到核算项目弹窗")

    keyword = counterparty_name if counterparty_name else fallback

    if not try_search_and_pick(page, dialog, keyword):
        if keyword != fallback:
            log(f"    ⚠ 未找到 [{keyword}]，改用兜底词：{fallback}")
            if not try_search_and_pick(page, dialog, fallback):
                close_project_dialog_if_open(page)
                raise RuntimeError(f"核算项目搜索失败：{keyword} / {fallback}")

    log("    等待弹窗自动关闭...")
    end = time.time() + 5
    closed = False
    while time.time() < end:
        if not is_project_dialog_open(page):
            log("    弹窗已关闭 ✅")
            closed = True
            break
        page.wait_for_timeout(300)

    if not closed:
        log("    ⚠ 5 秒内弹窗未关闭，继续下一步")

    page.wait_for_timeout(500)


def try_search_and_pick(page: Page, dialog: Locator, keyword: str) -> bool:
    log(f"    在搜索框输入：{keyword}")
    search_input = dialog.locator("input[type=text]").first
    search_input.wait_for(state="visible", timeout=8000)
    search_input.click()
    search_input.fill("")
    search_input.type(keyword, delay=40)
    page.wait_for_timeout(500)

    log("    点击放大镜按钮")
    clicked = False
    try:
        btn = dialog.locator("div.field-btn-right > button").first
        if btn.count() > 0:
            btn.click(timeout=3000)
            clicked = True
            log("    放大镜已点击（弹窗内）")
    except Exception as e:
        log(f"    弹窗内放大镜失败：{str(e)[:80]}")

    if not clicked:
        try:
            page.locator(Locators.MAGNIFIER).first.click(timeout=3000)
            clicked = True
            log("    放大镜已点击（全局）")
        except Exception as e:
            log(f"    全局放大镜失败：{str(e)[:80]}")

    if not clicked:
        log("    尝试按 Enter 搜索")
        page.keyboard.press("Enter")

    page.wait_for_timeout(2500)

    candidates = dialog.locator("span.tree-item-label").filter(has_text=keyword)
    cnt = candidates.count()
    log(f"    搜索结果数量（span.tree-item-label）：{cnt}")

    if cnt == 0:
        candidates = dialog.locator("div").filter(has_text=keyword)
        cnt = candidates.count()
        log(f"    搜索结果数量（div 兜底）：{cnt}")

    if cnt == 0:
        return False

    target = candidates.first
    for attempt in range(5):
        try:
            log(f"    尝试点击（第 {attempt + 1} 次）")
            target.click(timeout=2000)
            page.wait_for_timeout(1500)
            log("    ✅ 点击成功（弹窗应自动关闭）")
            return True
        except Exception as e:
            log(f"      失败：{str(e)[:80]}")
            target = target.locator("xpath=..")
    return False


# ========== 单条凭证 ==========
def enter_one_voucher(
    page: Page, row: dict[str, Any], index: int, total: int, bank_code: str
) -> None:
    voucher_no = row.get("voucher_no", "")
    no_txt = f"（凭证号 {voucher_no}）" if voucher_no else ""
    log("=" * 60)
    log(
        f"第 {index}/{total} 条{no_txt}：{row['direction']} | 日期 {row['date']} | "
        f"对方 {row['counterparty']} | 借方 {row['debit']:.2f} | 贷方 {row['credit']:.2f}"
    )
    log("=" * 60)

    # 先确保没有残留的核算项目弹窗
    close_project_dialog_if_open(page)
    page.wait_for_timeout(500)

    # 每条都新建独立凭证
    click(page, Locators.BTN_ADD, "增加（新建凭证）", wait=2500)

    # 填日期
    fill(page, Locators.DATE_INPUT, row["date"], "凭证日期", wait=500)
    page.keyboard.press("Tab")
    page.wait_for_timeout(500)

    # 录入凭证号：只有凭证号有文字时才填；为空则让 ERP 自动生成
    if voucher_no:
        fill(page, Locators.VOUCHER_NO_INPUT, voucher_no, "凭证号", wait=500)
        page.keyboard.press("Tab")
        page.wait_for_timeout(300)

    summary = row["summary"]

    # 付款
    if row["debit"] > 0:
        log("  ▶ 付款流程")
        amount = row["debit"]
        fill_summary(page, 1, summary)
        fill_and_confirm_subject(page, 1, SUBJECT_PAY_DEBIT)
        select_counterparty(page, row["counterparty"])
        fill_amount(page, 1, 7, amount, press_tab=True)
        fill_summary(page, 2, summary)
        fill_and_confirm_subject(page, 2, bank_code)
        fill_amount(page, 2, 8, amount, press_tab=False)

    # 收款
    elif row["credit"] > 0:
        log("  ▶ 收款流程")
        amount = row["credit"]
        fill_summary(page, 1, summary)
        fill_and_confirm_subject(page, 1, bank_code)
        fill_amount(page, 1, 7, amount, press_tab=True)
        fill_summary(page, 2, summary)
        fill_and_confirm_subject(page, 2, SUBJECT_RECEIVE_CREDIT)
        select_counterparty(page, row["counterparty"])
        fill_amount(page, 2, 8, amount, press_tab=False)

    else:
        log("  ⚠ 借贷金额都为 0，跳过")
        return

    # 保存
    click(page, Locators.BTN_SAVE, "保存凭证", wait=2500)
    log(f"  ✅ 第 {index} 条保存完成")

    # 关键：保存后关闭当前凭证录入窗口，准备下一条
    close_voucher_window(page)


def run_entry(
    excel_path: str | Path,
    bank_code: str | None = None,
    log_func: LogFunction | None = None,
) -> dict[str, Any]:
    """
    供 GUI 调用的入口函数。
    excel_path：已处理好的 Excel 文件路径。
    bank_code：银行科目码，可由 GUI 传入。
    log_func：可选日志回调。
    """
    if log_func is not None:
        global log

        def gui_log(msg):
            log_func(f"[{time.strftime('%H:%M:%S')}] {msg}")

        old_log = log
        log = gui_log

    try:
        if not os.path.exists(excel_path):
            raise FileNotFoundError(f"文件不存在：{excel_path}")

        bank_code = (bank_code or "").strip() or DEFAULT_BANK_CODE
        log(f"银行科目码：{bank_code}")

        log(f"读取 Excel：{excel_path}")
        rows = load_rows(excel_path)
        log(f"共 {len(rows)} 条待录入")
        entered = 0
        skipped = 0
        failed = 0

        with sync_playwright() as p:
            log(f"连接 Edge（{CDP_URL}）...")
            try:
                browser = p.chromium.connect_over_cdp(CDP_URL)
            except Exception as e:
                log(f"❌ 连接失败：{e}")
                log("请确认 FACAS 专用 Edge 正在运行，且 ERP 已完成登录")
                raise RuntimeError(f"无法连接 FACAS 专用 Edge：{e}") from e

            pages = [
                page
                for context in browser.contexts
                for page in context.pages
                if not page.is_closed() and "erp.bfcgj.com" in page.url
            ]
            if not pages:
                raise RuntimeError(
                    "未找到已登录的 ERP 页面，请先从 FACAS 点击‘登录 ERP’并完成登录"
                )
            page = pages[-1]
            context = page.context

            log("尝试导航到凭证录入（失败也没关系）...")
            try:
                click(page, Locators.MENU_FINANCE, "财务管理", wait=1500)
                click(page, Locators.MENU_VOUCHER, "凭证处理", wait=1500)
                click(page, Locators.MENU_ENTRY, "凭证录入", wait=2500)
            except Exception as e:
                log(f"⚠ 导航失败（继续）：{e}")

            for i, row in enumerate(rows, start=1):
                if str(row.get("skip", "")).strip() == "1":
                    skipped += 1
                    no_txt = (
                        f"（凭证号 {row.get('voucher_no', '')}）"
                        if row.get("voucher_no")
                        else ""
                    )
                    log(f"  ⏭ 跳过第 {i} 条{no_txt}：标记为「不需要录入凭证」")
                    continue
                try:
                    enter_one_voucher(page, row, i, len(rows), bank_code)
                    entered += 1
                except Exception as e:
                    failed += 1
                    log(f"  ❌ 第 {i} 条失败：{e}")
                    shot = str(Path(excel_path).resolve().parent / f"error_row_{i}.png")
                    try:
                        page.screenshot(path=shot, full_page=True)
                        log(f"  已截图：{shot}")
                    except Exception:
                        pass
                    close_project_dialog_if_open(page)
                    close_voucher_window(page)
                    page.wait_for_timeout(800)

            log("=" * 60)
            log(f"录入汇总：成功 {entered}，失败 {failed}，跳过 {skipped}")
            log("全部处理结束" if failed == 0 else "处理结束，存在失败凭证")
            log("=" * 60)
            return {
                "success": failed == 0,
                "entered": entered,
                "failed": failed,
                "skipped": skipped,
                "total": len(rows),
            }
    finally:
        if log_func is not None:
            log = old_log


def main() -> None:
    if len(sys.argv) < 2:
        print("用法：python auto_entry.py <Excel文件路径>")
        sys.exit(1)

    run_entry(sys.argv[1])


if __name__ == "__main__":
    main()
