# -*- coding: utf-8 -*-
"""银行流水处理、ERP 凭证录入及图形界面。

后续银行流水功能统一维护在本文件中。
"""

from __future__ import annotations

# ========== 银行流水 Excel 处理与付款单匹配 ==========

from pathlib import Path
from typing import Callable

import pandas as pd

LogFunction = Callable[[str], None]
ExcelPath = str | Path

# ========== 列名配置 ==========
# 按你的导入模板列名填写。如果以后模板列名变了，只需改这里，不用动下面的函数。
DATE_COL = "交易日期"  # 交易日期
SUMMARY_COL = "摘要"  # 摘要（会被自动生成的内容覆盖）
COUNTERPARTY_COL = "对方账户名称"
ACCOUNT_COL = "对方账号"
DEBIT_COL = "借方发生额"  # 有数字 → 付款
CREDIT_COL = "贷方发生额"  # 有数字 → 收款
REMARK_COL = "附言"  # 附言（银行流水里的附言栏）
DIRECTION_COL = "资金方向"  # 新增列，放在最后一列（H 栏）
VOUCHER_NO_COL = "凭证号"  # 新增列 I：凭证号（按文档顺序 1、2、3...）
SKIP_COL = "不需要录入凭证"  # 新增列 J：用户填 1 表示跳过这行，不录入


PAYMENT_BILL_NO_COL = "综合付款单号"
MATCH_STATUS_COL = "匹配状态"
AMOUNT_CHECK_COL = "金额校验"
PAYMENT_BILL_HEADERS = [
    "票据号码",
    "单据日期",
    "对方单位",
    "结算方式",
    "完成日期",
    "金额",
    "制单",
]

# 摘要中的方向描述词：资金方向列仍是"付"/"收"，这里只影响摘要文字
# 付 → 付残值款；收 → 收货款
DIRECTION_TEXT = {"付": "付残值款", "收": "收货款"}

# 读取时需要的列
REQUIRED_COLS = [
    DATE_COL,
    SUMMARY_COL,
    COUNTERPARTY_COL,
    ACCOUNT_COL,
    DEBIT_COL,
    CREDIT_COL,
    REMARK_COL,
]


def read_statement(file_path: ExcelPath, sheet_name: int | str = 0) -> pd.DataFrame:
    """读取银行流水 Excel，清洗金额和日期，返回标准格式的数据表。"""
    # 读取指定列，避免多余列干扰；列名对不上会给出明确提示
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, usecols=REQUIRED_COLS)
    except ValueError as e:
        # 读一下真实表头，告诉用户缺了哪些列
        actual = pd.read_excel(
            file_path, sheet_name=sheet_name, nrows=1
        ).columns.tolist()
        missing = [c for c in REQUIRED_COLS if c not in actual]
        raise ValueError(
            f"模板列名不匹配！\n你的模板列名: {actual}\n"
            f"缺少这些列: {missing}\n"
            f"请修改 bank_entry.py 顶部的列名配置（DATE_COL / DEBIT_COL 等）"
        ) from e

    # 清洗金额列：去掉千分位逗号，转数值型，空值/非法值统一置 0
    for col in [DEBIT_COL, CREDIT_COL]:
        df[col] = pd.to_numeric(
            df[col].astype(str).str.replace(",", "", regex=False), errors="coerce"
        ).fillna(0)

    # 日期列统一为字符串，方便打印和拼接摘要（解析失败则原样保留）
    parsed_date = pd.to_datetime(df[DATE_COL], errors="coerce")
    df[DATE_COL] = parsed_date.dt.strftime("%Y-%m-%d").fillna(df[DATE_COL].astype(str))

    return df


def _clean(value: object) -> str:
    """将 NaN/None 统一转为去掉首尾空格的字符串。"""
    if pd.isna(value):
        return ""
    return str(value).strip()


def validate_statement(df: pd.DataFrame) -> list[str]:
    """校验 Excel 内容是否合格。

    规则：交易日期不能为空；借方发生额和贷方发生额不能同时为 0。
    对方账户名称、对方账号可以为空。
    返回错误信息列表（空列表表示合格）。
    """
    errors = []
    for i, (_, row) in enumerate(df.iterrows(), start=1):
        date = _clean(row[DATE_COL])
        debit = float(row[DEBIT_COL]) if pd.notna(row[DEBIT_COL]) else 0.0
        credit = float(row[CREDIT_COL]) if pd.notna(row[CREDIT_COL]) else 0.0
        if not date:
            errors.append(f"第 {i} 行：交易日期为空")
        if debit == 0 and credit == 0:
            errors.append(
                f"第 {i} 行：借方发生额和贷方发生额都为 0（至少要有其中一个金额）"
            )
    return errors


def _read_optional_column(
    file_path: ExcelPath, sheet_name: int | str, col_name: str, nrows: int
) -> list[str]:
    """读取原表中的可选列；缺少该列时返回空字符串列表。"""
    try:
        header = pd.read_excel(
            file_path, sheet_name=sheet_name, nrows=0
        ).columns.tolist()
        if col_name in header:
            col = pd.read_excel(file_path, sheet_name=sheet_name, usecols=[col_name])
            values = col.iloc[:, 0].astype(str).replace(["nan", "None"], "").tolist()
            return (values + [""] * nrows)[:nrows]
    except Exception:
        pass
    return [""] * nrows


def preprocess_statement(df: pd.DataFrame, auto_summary: bool = True) -> pd.DataFrame:
    """执行流水预处理。

    借方发生额生成“付”方向，贷方发生额生成“收”方向。
    可选自动生成摘要；关闭时保留源表摘要。
    """
    # 1. 生成资金方向（始终生成）
    df[DIRECTION_COL] = [
        "付" if debit > 0 else ("收" if credit > 0 else "")
        for debit, credit in zip(df[DEBIT_COL], df[CREDIT_COL])
    ]

    # 2. 自动生成摘要（可选开关）
    if auto_summary:
        summaries = []
        for date, direction, counterparty, remark in zip(
            df[DATE_COL], df[DIRECTION_COL], df[COUNTERPARTY_COL], df[REMARK_COL]
        ):
            if not direction:
                # 没有借贷发生额的行，不生成摘要
                summaries.append("")
                continue
            parts = [
                _clean(date),
                DIRECTION_TEXT.get(direction, direction),
                _clean(counterparty),
                _clean(remark),
            ]
            summaries.append("-".join(p for p in parts if p))
        df[SUMMARY_COL] = summaries
    # auto_summary=False：保留原表摘要，不做任何修改

    return df


def save_processed(
    df: pd.DataFrame,
    output_path: ExcelPath = "processed_statement.xlsx",
    log_func: LogFunction | None = None,
) -> None:
    """将处理后的流水保存为新 Excel 文件，不覆盖原始流水。"""
    df.to_excel(output_path, index=False)
    message = f"已保存处理后的流水: {output_path}"
    (log_func or print)(message)


def process_statement(df: pd.DataFrame, log_func: LogFunction | None = None) -> None:
    """按借贷方向记录付款或收款摘要，可将日志输出到图形界面。"""
    emit = log_func or print
    for _, row in df.iterrows():
        date = row[DATE_COL]
        summary = "" if pd.isna(row[SUMMARY_COL]) else row[SUMMARY_COL]
        account = "" if pd.isna(row[COUNTERPARTY_COL]) else row[COUNTERPARTY_COL]
        debit = row[DEBIT_COL]
        credit = row[CREDIT_COL]

        # 借方有值且大于0 → 付款
        if debit > 0:
            emit(
                f"[付款] 日期:{date}, 摘要:{summary}, 对方账户:{account}, "
                f"第1行科目:220201, 第2行科目:10029334MZ, 金额:{debit:.2f}"
            )

        # 贷方有值且大于0 → 收款
        elif credit > 0:
            emit(
                f"[收款] 日期:{date}, 摘要:{summary}, 对方账户:{account}, "
                f"第1行科目:10029334MZ, 第2行科目:112201, 金额:{credit:.2f}"
            )

        # 两个金额均为空或为0 → 跳过该行


def process_bank_statement(
    file_path: ExcelPath,
    sheet_name: int | str = 0,
    output_path: ExcelPath = "processed_statement.xlsx",
    auto_summary: bool = True,
    log_func: LogFunction | None = None,
) -> None:
    """读取并校验银行流水，执行预处理后保存结果。"""
    df = read_statement(file_path, sheet_name)

    # 校验内容：有不合格就中止，不继续
    errors = validate_statement(df)
    if errors:
        raise ValueError("Excel 内容不合格，请修正后重新导入：\n" + "\n".join(errors))

    df = preprocess_statement(df, auto_summary=auto_summary)

    # 按表格行序生成凭证号。
    df[VOUCHER_NO_COL] = list(range(1, len(df) + 1))
    # 保留源表中的跳过标记；没有该列时补空值。
    skip_values = _read_optional_column(file_path, sheet_name, SKIP_COL, len(df))
    df[SKIP_COL] = skip_values

    save_processed(df, output_path, log_func=log_func)
    process_statement(df, log_func=log_func)


def _normalize_date(value: object) -> str:
    """将日期统一为 yyyy-mm-dd，无法解析时返回空字符串。"""
    if pd.isna(value):
        return ""
    parsed = pd.to_datetime(value, errors="coerce")
    if pd.isna(parsed):
        return ""
    return parsed.strftime("%Y-%m-%d")


def _normalize_company(value: object) -> str:
    """统一公司名称空白，避免空格差异影响匹配。"""
    if pd.isna(value):
        return ""
    return "".join(str(value).replace("\u3000", " ").split())


def _read_payment_bills(
    file_path: ExcelPath, sheet_name: int | str = 0
) -> pd.DataFrame:
    """读取 FACAS 导出的综合付款单，并校验所需字段。"""
    bills = pd.read_excel(file_path, sheet_name=sheet_name, dtype={"票据号码": str})
    missing = [header for header in PAYMENT_BILL_HEADERS if header not in bills.columns]
    if missing:
        raise ValueError(
            f"综合付款单 Excel 缺少字段：{missing}\n"
            f"需要的表头为：{PAYMENT_BILL_HEADERS}"
        )

    bills = bills[PAYMENT_BILL_HEADERS].copy()
    bills["票据号码"] = bills["票据号码"].map(_clean)
    bills["_日期键"] = bills["单据日期"].map(_normalize_date)
    bills["_公司键"] = bills["对方单位"].map(_normalize_company)
    bills["_金额"] = pd.to_numeric(
        bills["金额"].astype(str).str.replace(",", "", regex=False), errors="coerce"
    )

    # 不要在分组前丢弃单号或金额无效的记录：同日同公司下的无效记录
    # 仍可能是第二张付款单，过滤后会把多张单据误判为唯一匹配。
    return bills


def match_payment_bills(
    statement_path: ExcelPath,
    payment_bills_path: ExcelPath,
    output_path: ExcelPath,
    log_func: LogFunction | None = None,
) -> tuple[int, int, int, int]:
    """按付款日期和公司匹配综合付款单，校验流水合计后生成新 Excel。"""
    emit = log_func or print
    statement = pd.read_excel(statement_path, dtype={"凭证号": str})
    required = [DATE_COL, COUNTERPARTY_COL, DEBIT_COL, VOUCHER_NO_COL]
    missing = [column for column in required if column not in statement.columns]
    if missing:
        raise ValueError(f"银行流水 Excel 缺少字段：{missing}")

    bills = _read_payment_bills(payment_bills_path)
    statement[DEBIT_COL] = pd.to_numeric(
        statement[DEBIT_COL].astype(str).str.replace(",", "", regex=False),
        errors="coerce",
    ).fillna(0)
    statement["_日期键"] = statement[DATE_COL].map(_normalize_date)
    statement["_公司键"] = statement[COUNTERPARTY_COL].map(_normalize_company)

    # 显式使用统一后的字符串键建组，避免 pandas 对日期或公司值的类型推断造成键不一致。
    bill_groups = {}
    for _, bill in bills.iterrows():
        key = (str(bill["_日期键"]), str(bill["_公司键"]))
        bill_groups.setdefault(key, []).append(bill)

    statement_groups = {}
    for row_index, row in statement.iterrows():
        key = (str(row["_日期键"]), str(row["_公司键"]))
        if key[0] and key[1] and float(row[DEBIT_COL]) > 0:
            statement_groups.setdefault(key, []).append(row_index)

    statement[PAYMENT_BILL_NO_COL] = ""
    statement[MATCH_STATUS_COL] = "未匹配"
    statement[AMOUNT_CHECK_COL] = "-"

    for key, candidates in bill_groups.items():
        indices = statement_groups.get(key, [])
        if not key[0] or not key[1] or not indices:
            continue
        if len(candidates) != 1:
            bill_numbers = "、".join(
                bill["票据号码"] or "（单号为空）" for bill in candidates
            )
            statement.loc[indices, MATCH_STATUS_COL] = (
                f"待核对：同日同公司有多张付款单（{bill_numbers}）"
            )
            continue

        bill = candidates[0]
        bill_no = bill["票据号码"]
        if not bill_no or pd.isna(bill["_金额"]):
            statement.loc[indices, MATCH_STATUS_COL] = (
                "待核对：付款单单号或金额无效"
            )
            continue

        bill_total = round(float(bill["_金额"]), 2)
        flow_total = round(float(statement.loc[indices, DEBIT_COL].sum()), 2)
        # 日期和公司唯一对应一张单据时，整组付款流水都关联这张单据。
        # 先按流水顺序统一凭证号，再独立核对金额；不按金额猜测子组合。
        statement.loc[indices, PAYMENT_BILL_NO_COL] = bill_no
        first_voucher_id = statement.at[indices[0], VOUCHER_NO_COL]
        statement.loc[indices, VOUCHER_NO_COL] = first_voucher_id
        difference = round(bill_total - flow_total, 2)
        if difference == 0:
            statement.loc[indices, MATCH_STATUS_COL] = "已匹配"
            statement.loc[indices, AMOUNT_CHECK_COL] = "一致"
        else:
            statement.loc[indices, MATCH_STATUS_COL] = "已匹配"
            statement.loc[indices, AMOUNT_CHECK_COL] = (
                f"金额异常：差额 {difference:.2f}"
            )

    original_columns = pd.read_excel(statement_path, nrows=0).columns.tolist()
    for column in (PAYMENT_BILL_NO_COL, MATCH_STATUS_COL, AMOUNT_CHECK_COL):
        if column not in original_columns:
            original_columns.append(column)
    statement.drop(columns=["_日期键", "_公司键"], inplace=True)
    statement = statement[original_columns]
    statement.to_excel(output_path, index=False)
    matched = int((statement[MATCH_STATUS_COL] == "已匹配").sum())
    unchecked = int(statement[MATCH_STATUS_COL].str.startswith("待核对").sum())
    mismatched = int(statement[AMOUNT_CHECK_COL].str.startswith("金额异常").sum())
    emit(f"付款单匹配完成：流水 {len(statement)} 行，已匹配 {matched} 行")
    emit(f"待核对 {unchecked} 行，金额异常 {mismatched} 行")
    emit(f"匹配结果已保存：{output_path}")
    return len(statement), matched, unchecked, mismatched


# ========== ERP 凭证自动录入 ==========
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


# ========== 银行流水图形界面 ==========
import json
import os
import sys
import threading
import traceback
from pathlib import Path
from types import TracebackType
from typing import Any

from PySide6.QtCore import QObject, Qt, Signal
from PySide6.QtGui import QCloseEvent, QMouseEvent
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)



# ========== 常量 ==========
def _app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


APP_DIR = _app_dir()
CONFIG_FILE = APP_DIR / "bank-entry-settings.json"
LEGACY_CONFIG_FILE = APP_DIR / "jn" / "config.json"
ERROR_LOG = APP_DIR / "bank-entry-error.log"
SUMMARY_OUTPUT_NAME = "processed_statement.xlsx"
KEEP_OUTPUT_NAME = "processed_statement_keep.xlsx"
MATCHED_OUTPUT_NAME = "matched_payment_statement.xlsx"
WINDOW_TITLE = "银行流水凭证录入工具"


# ========== 主题样式 ==========
STYLE = """
QWidget#root {
    background: #edf4fc;
    color: #17324d;
}
QFrame#headerCard {
    background: #0a2948;
    border: 1px solid #164d78;
    border-radius: 18px;
}
QFrame#accentBar {
    background: #44c5f4;
    border-radius: 3px;
}
QLabel#eyebrow {
    color: #8fc7e8;
    font-size: 10px;
    font-weight: 700;
    letter-spacing: 1px;
}
QLabel#pageTitle {
    color: #ffffff;
    font-size: 22px;
    font-weight: 700;
}
QLabel#pageSubtitle {
    color: #b9d5ea;
    font-size: 11px;
}
QFrame#card {
    background: #ffffff;
    border: 1px solid #cfe0f2;
    border-radius: 14px;
}
QFrame#actionBar {
    background: #f7fbff;
    border: 1px solid #c5dced;
    border-radius: 14px;
}
QLabel#cardTitle {
    color: #17324d;
    font-size: 12px;
    font-weight: 700;
}
QLabel#fieldLabel {
    color: #54708a;
    font-size: 12px;
    font-weight: 600;
}
QLineEdit {
    background: #ffffff;
    border: 1px solid #bcd2e8;
    border-radius: 10px;
    padding: 7px 10px;
    min-height: 21px;
    color: #17324d;
    font-size: 12px;
    selection-background-color: #b8e4fb;
}
QLineEdit:focus {
    border: 1px solid #2589c7;
}
QCheckBox {
    color: #294863;
    spacing: 8px;
    padding: 3px 0;
    font-size: 12px;
}
QCheckBox:checked {
    color: #1261a0;
    font-weight: 700;
}
QCheckBox::indicator {
    width: 17px;
    height: 17px;
    border: 1px solid #a9c4dc;
    border-radius: 5px;
    background: #ffffff;
}
QCheckBox::indicator:hover {
    border: 1px solid #2589c7;
}
QCheckBox::indicator:checked {
    background: #1676b8;
    border: 1px solid #1676b8;
}
QPushButton {
    border: 1px solid #bcd2e8;
    border-radius: 8px;
    background: #ffffff;
    color: #2e587b;
    padding: 7px 12px;
    font-weight: 600;
    font-size: 11px;
}
QPushButton:hover {
    background: #edf7ff;
    border-color: #75b8df;
}
QPushButton:pressed {
    background: #d9effc;
}
QPushButton:disabled {
    background: #eaf1f8;
    color: #91a5b8;
    border-color: #d9e5f0;
}
QPushButton#primaryButton {
    background: #1267a8;
    color: #ffffff;
    border: none;
    border-radius: 10px;
    padding: 10px 24px;
    font-size: 11px;
    font-weight: 700;
}
QPushButton#primaryButton:hover {
    background: #0e5590;
}
QPushButton#primaryButton:pressed {
    background: #0a416e;
}
QPushButton#primaryButton:disabled {
    background: #86b8d7;
    color: #eaf7ff;
}
QPushButton#autoButton {
    background: #0f9e6c;
    color: #ffffff;
    border: none;
    border-radius: 10px;
    padding: 10px 24px;
    font-size: 11px;
    font-weight: 700;
}
QPushButton#autoButton:hover {
    background: #0c8459;
}
QPushButton#autoButton:pressed {
    background: #0a6e4a;
}
QPushButton#autoButton:disabled {
    background: #7ec4a8;
    color: #eaf7ff;
}
QPushButton#secondaryButton {
    background: #e5f2fc;
    color: #1261a0;
    border: 1px solid #9fc8e5;
    border-radius: 10px;
    padding: 10px 18px;
    font-size: 12px;
    font-weight: 700;
}
QPushButton#secondaryButton:hover {
    background: #d7edfb;
    border-color: #62a9d2;
}
QPushButton#secondaryButton:pressed {
    background: #c6e5f7;
}
QTextEdit#logView {
    background: #0c2137;
    color: #c9eeff;
    border: 1px solid #17496d;
    border-radius: 10px;
    padding: 10px;
    selection-background-color: #15527b;
    font-family: "Consolas", "Microsoft YaHei";
    font-size: 11px;
}
QScrollBar:vertical {
    background: #e6f0fa;
    width: 12px;
    border: 1px solid #c7dceb;
    border-radius: 6px;
}
QScrollBar::handle:vertical {
    background: #8cb9d8;
    min-height: 30px;
    border-radius: 4px;
}
QScrollBar::handle:vertical:hover {
    background: #5d9bc4;
}
QScrollBar::add-line:vertical,
QScrollBar::sub-line:vertical {
    height: 0px;
}
"""


# ========== 配置文件读写 ==========
def load_config() -> dict[str, Any]:
    config_path = CONFIG_FILE if CONFIG_FILE.exists() else LEGACY_CONFIG_FILE
    if not config_path.exists():
        return {}
    try:
        with open(config_path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_config(config: dict[str, Any]) -> None:
    try:
        CONFIG_FILE.parent.mkdir(parents=True, exist_ok=True)
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"[警告] 保存配置文件失败: {e}")


def write_error_log(text: str) -> None:
    try:
        with open(ERROR_LOG, "a", encoding="utf-8") as f:
            f.write(text + "\n")
    except Exception:
        pass


def _excepthook(
    exc_type: type[BaseException], exc: BaseException, tb: TracebackType | None
) -> None:
    write_error_log("".join(traceback.format_exception(exc_type, exc, tb)))
    try:
        if QApplication.instance() is not None:
            QMessageBox.critical(None, "程序出错", f"发生错误：\n{exc}")
    except Exception:
        pass


# ========== 防覆盖文件名生成 ==========
def get_non_exist_path(folder: str | Path, filename: str) -> str:
    base, ext = os.path.splitext(filename)
    candidate = os.path.join(folder, filename)
    if not os.path.exists(candidate):
        return candidate
    i = 1
    while True:
        candidate = os.path.join(folder, f"{base}({i}){ext}")
        if not os.path.exists(candidate):
            return candidate
        i += 1


# ========== 后台线程 → 界面日志的安全桥（跨线程用信号，避免 Qt 崩溃） ==========
class LogBridge(QObject):
    message = Signal(str)
    finished = Signal(object)


# ========== 主窗口 ==========
class MainWindow(QWidget):
    bank_task_started = Signal()
    bank_task_finished = Signal(str)

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._external_busy = False
        self._bank_operation_running = False
        self.setObjectName("root")
        self.setWindowTitle(WINDOW_TITLE)
        self.resize(760, 640)
        # 明确使用普通顶层窗体标志，确保 Windows 将最小化按钮识别为真正的最小化。
        self.setWindowFlags(
            Qt.WindowType.Window
            | Qt.WindowType.WindowTitleHint
            | Qt.WindowType.WindowSystemMenuHint
            | Qt.WindowType.WindowMinimizeButtonHint
            | Qt.WindowType.WindowCloseButtonHint
        )

        self.config = load_config()
        self.output_dir = self.config.get("output_dir", "")
        self.user_chose_output = bool(self.output_dir)
        self.source_file = ""
        self.payment_bills_file = ""
        self.processed_file = None
        self.bank_code = self.config.get("bank_code", "10029334MZ")
        self.payment_bills_file = self.config.get("payment_bills_file", "")

        # 供 auto_entry 的日志和"停止"标志使用
        self._stop_flag = False
        self.log_bridge = LogBridge()
        self.log_bridge.message.connect(self.append_log)
        self.log_bridge.finished.connect(self.on_entry_finished)

        self._build_ui()
        self.setStyleSheet(STYLE)
        self._refresh_action_controls()

        if self.output_dir:
            self.output_path_edit.setText(self.output_dir)

    def _build_ui(self) -> None:
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(18, 18, 18, 18)
        main_layout.setSpacing(12)

        # ---- 顶部深蓝卡片 ----
        header = QFrame()
        header.setObjectName("headerCard")
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(20, 18, 20, 18)
        header_layout.setSpacing(14)
        accent = QFrame()
        accent.setObjectName("accentBar")
        accent.setFixedWidth(5)
        accent.setMinimumHeight(68)
        header_layout.addWidget(accent)

        title_box = QVBoxLayout()
        title_box.setSpacing(4)
        eyebrow = QLabel("银行流水 · 凭证处理")
        eyebrow.setObjectName("eyebrow")
        title_box.addWidget(eyebrow)

        title = QLabel("银行流水凭证录入工具")
        title.setObjectName("pageTitle")
        title_box.addWidget(title)

        subtitle = QLabel("读取 Excel · 判断方向 · 生成摘要 · 自动录入 ERP")
        subtitle.setObjectName("pageSubtitle")
        title_box.addWidget(subtitle)
        header_layout.addLayout(title_box)
        header_layout.addStretch(1)
        main_layout.addWidget(header)

        # ---- 源文件卡片 ----
        source_card = QFrame()
        source_card.setObjectName("card")
        source_layout = QVBoxLayout(source_card)
        source_layout.setContentsMargins(14, 12, 14, 12)
        source_layout.setSpacing(8)

        source_title = QLabel("源文件")
        source_title.setObjectName("cardTitle")
        source_layout.addWidget(source_title)

        row1 = QHBoxLayout()
        self.choose_btn = QPushButton("选择Excel文件")
        self.choose_btn.clicked.connect(self.on_choose_file)
        self.source_path_edit = QLineEdit()
        self.source_path_edit.setReadOnly(True)
        self.source_path_edit.setPlaceholderText("尚未选择文件")
        row1.addWidget(self.choose_btn)
        row1.addWidget(self.source_path_edit, 1)
        source_layout.addLayout(row1)

        payment_row = QHBoxLayout()
        self.choose_payment_bills_btn = QPushButton("选择综合付款单 Excel")
        self.choose_payment_bills_btn.clicked.connect(self.on_choose_payment_bills)
        self.payment_bills_path_edit = QLineEdit()
        self.payment_bills_path_edit.setReadOnly(True)
        self.payment_bills_path_edit.setPlaceholderText(
            "选择从 FACAS 导出的综合付款单 Excel"
        )
        self.payment_bills_path_edit.setText(self.payment_bills_file)
        payment_row.addWidget(self.choose_payment_bills_btn)
        payment_row.addWidget(self.payment_bills_path_edit, 1)
        source_layout.addLayout(payment_row)
        main_layout.addWidget(source_card)

        # ---- 处理选项卡片 ----
        options_card = QFrame()
        options_card.setObjectName("card")
        options_layout = QVBoxLayout(options_card)
        options_layout.setContentsMargins(14, 12, 14, 12)
        options_layout.setSpacing(8)

        options_title = QLabel("处理选项")
        options_title.setObjectName("cardTitle")
        options_layout.addWidget(options_title)

        # 一行：自动生成摘要 + 输出位置 + 输出文件夹显示框 + 打开输出文件夹
        option_row = QHBoxLayout()
        self.auto_summary_cb = QCheckBox("自动生成摘要")
        self.auto_summary_cb.setChecked(True)
        option_row.addWidget(self.auto_summary_cb)

        output_label = QLabel("输出位置")
        output_label.setObjectName("fieldLabel")
        option_row.addWidget(output_label)

        self.output_path_edit = QLineEdit()
        self.output_path_edit.setReadOnly(True)
        self.output_path_edit.setPlaceholderText("默认跟随源文件所在文件夹")
        self.output_path_edit.setCursor(Qt.PointingHandCursor)
        self.output_path_edit.mousePressEvent = self.on_click_output_edit
        option_row.addWidget(self.output_path_edit, 1)

        self.open_btn = QPushButton("打开输出文件夹")
        self.open_btn.setObjectName("secondaryButton")
        self.open_btn.setCursor(Qt.PointingHandCursor)
        self.open_btn.clicked.connect(self.on_open_output)
        option_row.addWidget(self.open_btn)
        options_layout.addLayout(option_row)

        # 银行科目码（会记忆，下次打开还是这个值）
        bank_row = QHBoxLayout()
        bank_label = QLabel("银行科目码")
        bank_label.setObjectName("fieldLabel")
        bank_label.setFixedWidth(70)
        self.bank_code_edit = QLineEdit()
        self.bank_code_edit.setPlaceholderText("例如：10029334MZ")
        self.bank_code_edit.setText(self.bank_code)
        self.bank_code_edit.editingFinished.connect(self.on_bank_code_changed)
        bank_row.addWidget(bank_label)
        bank_row.addWidget(self.bank_code_edit, 1)
        options_layout.addLayout(bank_row)

        main_layout.addWidget(options_card)

        # ---- 操作栏 ----
        action_bar = QFrame()
        action_bar.setObjectName("actionBar")
        action_layout = QHBoxLayout(action_bar)
        action_layout.setContentsMargins(14, 12, 14, 12)
        action_layout.setSpacing(10)

        self.run_btn = QPushButton("① 开始处理（生成中间Excel）")
        self.run_btn.setObjectName("primaryButton")
        self.run_btn.setCursor(Qt.PointingHandCursor)
        self.run_btn.clicked.connect(self.on_run)
        action_layout.addWidget(self.run_btn, 1)

        self.entry_btn = QPushButton("② 开始录入 ERP")
        self.entry_btn.setObjectName("autoButton")
        self.entry_btn.setCursor(Qt.PointingHandCursor)
        self.entry_btn.clicked.connect(self.on_entry)
        action_layout.addWidget(self.entry_btn, 1)

        self.match_btn = QPushButton("③ 匹配综合付款单")
        self.match_btn.setObjectName("secondaryButton")
        self.match_btn.setCursor(Qt.PointingHandCursor)
        self.match_btn.clicked.connect(self.on_match_payment_bills)
        action_layout.addWidget(self.match_btn, 1)

        main_layout.addWidget(action_bar)

        # ---- 日志卡片 ----
        log_card = QFrame()
        log_card.setObjectName("card")
        log_layout = QVBoxLayout(log_card)
        log_layout.setContentsMargins(14, 12, 14, 12)
        log_layout.setSpacing(8)

        log_title = QLabel("日志输出")
        log_title.setObjectName("cardTitle")
        log_layout.addWidget(log_title)

        self.log_edit = QTextEdit()
        self.log_edit.setObjectName("logView")
        self.log_edit.setReadOnly(True)
        log_layout.addWidget(self.log_edit, 1)
        main_layout.addWidget(log_card, 1)

        # 处理日志通过显式回调和 Qt 信号传递，不覆盖宿主 FACAS 的标准输出。

    def set_external_busy(self, busy: bool) -> None:
        """与宿主 FACAS 共用 ERP 时，阻止并行运行另一项浏览器任务。"""
        self._external_busy = bool(busy)
        self._refresh_action_controls()

    def _refresh_action_controls(self) -> None:
        available = not self._external_busy and not self._bank_operation_running
        if hasattr(self, "run_btn"):
            self.run_btn.setEnabled(available)
            self.entry_btn.setEnabled(available)
            self.choose_payment_bills_btn.setEnabled(available)
            self.match_btn.setEnabled(available)

    # ---------- 事件 ----------
    def on_choose_file(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择银行流水 Excel 文件",
            str(APP_DIR),
            "Excel 文件 (*.xlsx *.xls)",
        )
        if not file_path:
            return
        self.source_file = file_path
        self.processed_file = None
        self.source_path_edit.setText(file_path)
        if not self.user_chose_output:
            self.output_dir = os.path.dirname(file_path)
            self.output_path_edit.setText(self.output_dir)

    def on_choose_payment_bills(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择综合付款单 Excel 文件",
            str(APP_DIR),
            "Excel 文件 (*.xlsx *.xls)",
        )
        if not file_path:
            return
        self.payment_bills_file = file_path
        self.payment_bills_path_edit.setText(file_path)
        self.config["payment_bills_file"] = file_path
        save_config(self.config)
        self._refresh_action_controls()

    def on_match_payment_bills(self) -> None:
        if not self.source_file or not os.path.isfile(self.source_file):
            QMessageBox.warning(self, "提示", "请先选择有效的银行流水 Excel 文件。")
            return
        statement_file = self._find_latest_processed() or self.source_file
        if not self.payment_bills_file or not os.path.isfile(self.payment_bills_file):
            QMessageBox.warning(self, "提示", "请先选择有效的综合付款单 Excel 文件。")
            return
        if os.path.abspath(statement_file) == os.path.abspath(self.payment_bills_file):
            QMessageBox.warning(
                self, "提示", "银行流水和综合付款单不能选择同一个文件。"
            )
            return

        output_path = get_non_exist_path(self.output_dir, MATCHED_OUTPUT_NAME)
        self._bank_operation_running = True
        self._refresh_action_controls()
        self.bank_task_started.emit()
        self.match_btn.setText("匹配中...")
        QApplication.processEvents()
        succeeded = False

        try:
            self.append_log("=" * 60)
            self.append_log(f"开始匹配银行流水：{statement_file}")
            self.append_log(f"综合付款单：{self.payment_bills_file}")
            self.append_log(f"输出文件：{output_path}")
            result = match_payment_bills(
                statement_file,
                self.payment_bills_file,
                output_path,
                log_func=self.append_log,
            )
            succeeded = True
            QMessageBox.information(
                self,
                "匹配完成",
                f"已生成匹配结果：\n{output_path}\n\n"
                f"流水 {result[0]} 行，已匹配 {result[1]} 行，"
                f"待核对 {result[2]} 行，金额异常 {result[3]} 行。",
            )
        except Exception as e:
            self.append_log(f"付款单匹配失败：{e}")
            QMessageBox.critical(self, "匹配失败", f"处理过程中出错：\n\n{e}")
        finally:
            self.match_btn.setText("③ 匹配综合付款单")
            self._bank_operation_running = False
            self._refresh_action_controls()
            self.bank_task_finished.emit("success" if succeeded else "failure")

    def on_click_output_edit(self, event: QMouseEvent) -> None:
        folder = QFileDialog.getExistingDirectory(
            self,
            "选择输出文件夹",
            self.output_dir or str(APP_DIR),
        )
        if not folder:
            return
        self.output_dir = folder
        self.processed_file = None
        self.user_chose_output = True
        self.output_path_edit.setText(folder)
        self.config["output_dir"] = folder
        save_config(self.config)

    def on_open_output(self) -> None:
        if self.output_dir and os.path.isdir(self.output_dir):
            os.startfile(self.output_dir)
        else:
            QMessageBox.information(self, "提示", "尚未选择输出文件夹。")

    def on_bank_code_changed(self) -> None:
        self.bank_code = self.bank_code_edit.text().strip()
        self.config["bank_code"] = self.bank_code
        save_config(self.config)

    # ---------- ① 处理 Excel ----------
    def on_run(self) -> None:
        if not self._check_before_run():
            return

        self.processed_file = None
        succeeded = False
        auto_summary = self.auto_summary_cb.isChecked()
        base_name = SUMMARY_OUTPUT_NAME if auto_summary else KEEP_OUTPUT_NAME
        output_path = get_non_exist_path(self.output_dir, base_name)

        self._bank_operation_running = True
        self._refresh_action_controls()
        self.bank_task_started.emit()
        self.run_btn.setText("处理中...")
        QApplication.processEvents()

        try:
            self.append_log("=" * 60)
            self.append_log(f"开始处理：{self.source_file}")
            self.append_log(f"自动生成摘要：{'是' if auto_summary else '否'}")
            self.append_log(f"输出文件：{output_path}")
            self.append_log("=" * 60)

            process_bank_statement(
                self.source_file,
                output_path=output_path,
                auto_summary=auto_summary,
                log_func=self.append_log,
            )
            self.processed_file = output_path
            succeeded = True

            self.append_log("=" * 60)
            self.append_log(f"✅ 处理完成！文件已保存到：\n{output_path}")
            self.append_log("=" * 60)
            QMessageBox.information(
                self,
                "处理完成",
                f"处理完成！\n\n文件已保存到：\n{output_path}\n\n"
                f"接下来可以点击「② 开始录入 ERP」进行自动录入。",
            )
        except Exception as e:
            self.append_log(f"❌ 出错：{e}")
            QMessageBox.critical(self, "处理失败", f"处理过程中出错：\n\n{e}")
        finally:
            self.run_btn.setText("① 开始处理（生成中间Excel）")
            self._bank_operation_running = False
            self._refresh_action_controls()
            self.bank_task_finished.emit("success" if succeeded else "failure")

    # ---------- ② 录入 ERP ----------
    def on_entry(self) -> None:
        if not self._check_before_run():
            return

        target = self.processed_file
        if not target or not os.path.isfile(target):
            self.processed_file = None
            QMessageBox.warning(
                self,
                "提示",
                "本窗口还没有成功生成可录入的中间文件。\n\n"
                "请先点击「① 开始处理」生成中间文件。",
            )
            return

        reply = QMessageBox.question(
            self,
            "确认录入",
            f"即将读取以下文件并录入 ERP：\n\n{target}\n\n"
            f"请确认已用调试模式启动 Edge 并登录 ERP。\n\n是否继续？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if reply != QMessageBox.Yes:
            return

        self._bank_operation_running = True
        self._refresh_action_controls()
        self.bank_task_started.emit()
        self.entry_btn.setText("录入中...")
        QApplication.processEvents()

        bank_code = self.bank_code_edit.text().strip() or "10029334MZ"
        self.config["bank_code"] = bank_code
        save_config(self.config)

        # 用后台线程跑，避免界面卡死
        def worker() -> None:
            try:
                result = run_entry(
                    target, bank_code=bank_code, log_func=self.log_bridge.message.emit
                )
            except Exception as e:
                self.log_bridge.message.emit(f"❌ 录入失败：{e}")
                write_error_log(traceback.format_exc())
                result = {"success": False, "failed": 1, "error": str(e)}
            self.log_bridge.finished.emit(result)

        threading.Thread(target=worker, daemon=True).start()

    def on_entry_finished(self, result: Any) -> None:
        self._bank_operation_running = False
        self._refresh_action_controls()
        self.entry_btn.setText("② 开始录入 ERP")
        if not isinstance(result, dict) or not result.get("success"):
            outcome = (
                "partial"
                if isinstance(result, dict) and result.get("entered", 0)
                else "failure"
            )
        else:
            outcome = "success"
        self.bank_task_finished.emit(outcome)
        entered = result.get("entered", 0) if isinstance(result, dict) else 0
        failed = result.get("failed", 0) if isinstance(result, dict) else 1
        skipped = result.get("skipped", 0) if isinstance(result, dict) else 0
        self.append_log("=" * 60)
        self.append_log(f"录入结果：成功 {entered}，失败 {failed}，跳过 {skipped}")
        self.append_log("=" * 60)

    def _check_before_run(self) -> bool:
        if not self.source_file:
            QMessageBox.warning(self, "提示", "请先选择源 Excel 文件。")
            return False
        if not os.path.exists(self.source_file):
            QMessageBox.warning(self, "提示", "源文件不存在，请重新选择。")
            return False
        if not self.output_dir or not os.path.isdir(self.output_dir):
            QMessageBox.warning(self, "提示", "请先选择有效的输出位置。")
            return False
        return True

    def _find_latest_processed(self) -> str | None:
        """只返回本窗口成功生成的中间文件，避免误用目录中的旧文件。"""
        if not self.processed_file or not os.path.isfile(self.processed_file):
            return None
        return self.processed_file

    def append_log(self, text: object) -> None:
        text = str(text).rstrip("\n")
        if not text:
            return
        self.log_edit.append(text)
        self.log_edit.verticalScrollBar().setValue(
            self.log_edit.verticalScrollBar().maximum()
        )

    def closeEvent(self, event: QCloseEvent) -> None:
        if self._bank_operation_running:
            QMessageBox.warning(
                self, "任务进行中", "银行流水任务仍在运行，请等待完成后再关闭窗口。"
            )
            event.ignore()
            return
        # 关窗口时把银行科目码等配置保存好
        try:
            self.config["bank_code"] = self.bank_code_edit.text().strip()
            save_config(self.config)
        except Exception:
            pass
        event.accept()


def main() -> int:
    # Playwright 连接自检：验证 exe 内驱动能连上调试端口（正常使用不受影响）
    if len(sys.argv) >= 2 and sys.argv[1] == "--pwtest":
        try:
            from playwright.sync_api import sync_playwright

            with sync_playwright() as p:
                browser = p.chromium.connect_over_cdp("http://localhost:9222")
                pages = browser.contexts[0].pages if browser.contexts else []
                with open(APP_DIR / "pwtest.log", "w", encoding="utf-8") as f:
                    f.write(f"PW_CONNECT_OK pages={len(pages)}\n")
                return 0
        except Exception as e:
            with open(APP_DIR / "pwtest.log", "w", encoding="utf-8") as f:
                f.write(f"PW_CONNECT_FAIL {e}\n")
            return 1

    # 自检模式：用于打包后验证 pandas/openpyxl 是否被正确打包
    if len(sys.argv) >= 4 and sys.argv[1] == "--selftest":
        src, out = sys.argv[2], sys.argv[3]
        process_bank_statement(src, output_path=out, auto_summary=True)
        return 0

    sys.excepthook = _excepthook
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main() or 0)
