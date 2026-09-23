import tempfile
import unittest
from datetime import datetime
from pathlib import Path

import pandas as pd

from jn.process_bank import match_payment_bills


class PaymentMatchingTests(unittest.TestCase):
    def test_invalid_bill_rows_block_unsafe_auto_matching(self):
        statement = pd.DataFrame(
            {
                "交易日期": ["2026-09-07", "2026-09-08"],
                "摘要": ["", ""],
                "对方账户名称": ["己公司", "庚公司"],
                "对方账号": ["", ""],
                "借方发生额": [20, 30],
                "贷方发生额": [0, 0],
                "附言": ["", ""],
                "凭证号": ["1300", "1301"],
            }
        )
        bills = pd.DataFrame(
            [
                ["ZF007", datetime(2026, 9, 7), "己公司", "转账", None, 20, "制单人"],
                ["ZF008", datetime(2026, 9, 7), "己公司", "转账", None, None, "制单人"],
                ["", datetime(2026, 9, 8), "庚公司", "转账", None, 30, "制单人"],
            ],
            columns=[
                "票据号码",
                "单据日期",
                "对方单位",
                "结算方式",
                "完成日期",
                "金额",
                "制单",
            ],
        )

        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            statement_path = root / "statement.xlsx"
            bills_path = root / "bills.xlsx"
            output_path = root / "matched.xlsx"
            statement.to_excel(statement_path, index=False)
            bills.to_excel(bills_path, index=False)

            match_payment_bills(statement_path, bills_path, output_path)
            output = pd.read_excel(output_path, dtype={"凭证号": str})

        # 无效付款单仍参与日期+公司的唯一性判断，不能错误关联或合并凭证号。
        self.assertTrue(
            str(output.loc[0, "匹配状态"]).startswith("待核对：同日同公司有多张付款单")
        )
        self.assertTrue(pd.isna(output.loc[0, "综合付款单号"]))
        self.assertEqual(output.loc[0, "凭证号"], "1300")

        # 唯一但缺少票据号码的记录也不能被当作有效付款单。
        self.assertEqual(output.loc[1, "匹配状态"], "待核对：付款单单号或金额无效")
        self.assertTrue(pd.isna(output.loc[1, "综合付款单号"]))
        self.assertEqual(output.loc[1, "凭证号"], "1301")

    def test_matches_by_date_and_company_and_checks_whole_group(self):
        records = [
            ("2026-09-01", "甲公司", 100, "1250"),
            (datetime(2026, 9, 1), " 甲 公司 ", 200, "1251"),
            ("2026-09-02", "乙公司", 50, "1252"),
            ("2026-09-03", "丙公司", 80, "1253"),
            ("2026-09-04", "丁公司", 70, "1254"),
            ("2026-09-04", "丁公司", 30, "1255"),
            ("2026-09-04", "丁公司", 100, "1256"),
            ("2026-09-05", "甲公司", 25, "1257"),
            ("2026-09-06", "戊公司", 40, "1258"),
        ]
        statement = pd.DataFrame(
            {
                "交易日期": [row[0] for row in records],
                "摘要": [""] * len(records),
                "对方账户名称": [row[1] for row in records],
                "对方账号": [""] * len(records),
                "借方发生额": [row[2] for row in records],
                "贷方发生额": [0] * len(records),
                "附言": [""] * len(records),
                "凭证号": [row[3] for row in records],
            }
        )
        bill_rows = [
            [
                "ZF001",
                datetime(2026, 9, 1),
                "甲公司",
                "转账",
                datetime(2026, 9, 1),
                300,
                "制单人",
            ],
            [
                "ZF002",
                datetime(2026, 9, 2),
                "乙公司",
                "转账",
                datetime(2026, 9, 2),
                60,
                "制单人",
            ],
            [
                "ZF003",
                datetime(2026, 9, 3),
                "丙公司",
                "转账",
                datetime(2026, 9, 3),
                80,
                "制单人",
            ],
            [
                "ZF004",
                datetime(2026, 9, 3),
                "丙公司",
                "转账",
                datetime(2026, 9, 3),
                80,
                "制单人",
            ],
            [
                "ZF005",
                datetime(2026, 9, 4),
                "丁公司",
                "转账",
                datetime(2026, 9, 4),
                100,
                "制单人",
            ],
            [
                "ZF006",
                datetime(2026, 9, 6),
                "戊公司",
                "转账",
                datetime(2026, 9, 6),
                40,
                "制单人",
            ],
        ]
        bills = pd.DataFrame(
            bill_rows,
            columns=[
                "票据号码",
                "单据日期",
                "对方单位",
                "结算方式",
                "完成日期",
                "金额",
                "制单",
            ],
        )

        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            statement_path = root / "statement.xlsx"
            bills_path = root / "bills.xlsx"
            output_path = root / "matched.xlsx"
            statement.to_excel(statement_path, index=False)
            bills.to_excel(bills_path, index=False)

            result = match_payment_bills(statement_path, bills_path, output_path)
            output = pd.read_excel(output_path, dtype={"凭证号": str})

        # 多对一金额相符时，统一为 Excel 顺序第一条流水的凭证号。
        self.assertEqual(output.loc[0:1, "综合付款单号"].tolist(), ["ZF001", "ZF001"])
        self.assertEqual(output.loc[0:1, "凭证号"].tolist(), ["1250", "1250"])
        self.assertEqual(output.loc[0:1, "金额校验"].tolist(), ["一致", "一致"])

        # 金额异常仍保留日期和公司匹配结果，并按同一付款单号归并凭证号。
        self.assertEqual(output.loc[2, "综合付款单号"], "ZF002")
        self.assertEqual(output.loc[2, "凭证号"], "1252")
        self.assertEqual(output.loc[2, "金额校验"], "金额异常：差额 10.00")
        self.assertEqual(output.loc[4:6, "综合付款单号"].tolist(), ["ZF005"] * 3)
        self.assertEqual(output.loc[4:6, "凭证号"].tolist(), ["1254"] * 3)
        self.assertEqual(
            output.loc[4:6, "金额校验"].tolist(), ["金额异常：差额 -100.00"] * 3
        )

        # 同日同公司多张单据不猜测；日期不同也不能串到同一付款单。
        self.assertTrue(pd.isna(output.loc[3, "综合付款单号"]))
        self.assertTrue(
            str(output.loc[3, "匹配状态"]).startswith("待核对：同日同公司有多张付款单")
        )
        self.assertEqual(output.loc[3, "凭证号"], "1253")
        self.assertTrue(pd.isna(output.loc[7, "综合付款单号"]))
        self.assertEqual(output.loc[7, "匹配状态"], "未匹配")

        self.assertEqual(output.loc[8, "综合付款单号"], "ZF006")
        self.assertEqual(result, (9, 7, 1, 4))


if __name__ == "__main__":
    unittest.main()
