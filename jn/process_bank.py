from __future__ import annotations

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
            f"请修改 process_bank.py 顶部的列名配置（DATE_COL / DEBIT_COL 等）"
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


if __name__ == "__main__":
    # 独立运行示例：替换成与脚本同目录下的流水文件名。
    FILE = "银行对账单导入模板(回.xlsx"

    print("========== 自动生成摘要 ==========")
    process_bank_statement(FILE, output_path="processed_statement.xlsx")

    print("========== 保留源表摘要 ==========")
    process_bank_statement(
        FILE, output_path="processed_statement_keep.xlsx", auto_summary=False
    )
