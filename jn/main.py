# -*- coding: utf-8 -*-
"""
银行流水凭证自动录入工具 - GUI 主程序
依赖：pandas, openpyxl, PySide6, playwright
"""

from __future__ import annotations

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

if __package__:
    from . import process_bank, auto_entry
else:
    import process_bank
    import auto_entry


# ========== 常量 ==========
def _app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


APP_DIR = _app_dir()
CONFIG_FILE = APP_DIR / "config.json"
ERROR_LOG = APP_DIR / "error.log"
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
    if not CONFIG_FILE.exists():
        return {}
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_config(config: dict[str, Any]) -> None:
    try:
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
        self.setWindowFlag(Qt.WindowType.Window, True)

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
        header_layout = QVBoxLayout(header)
        header_layout.setContentsMargins(18, 16, 18, 16)
        header_layout.setSpacing(6)

        top_row = QHBoxLayout()
        accent = QFrame()
        accent.setObjectName("accentBar")
        accent.setFixedSize(26, 3)
        eyebrow = QLabel("银行流水 · 凭证处理")
        eyebrow.setObjectName("eyebrow")
        top_row.addWidget(accent)
        top_row.addWidget(eyebrow)
        top_row.addStretch(1)
        header_layout.addLayout(top_row)

        title = QLabel("银行流水凭证录入工具")
        title.setObjectName("pageTitle")
        header_layout.addWidget(title)

        subtitle = QLabel("读取 Excel · 判断方向 · 生成摘要 · 自动录入 ERP")
        subtitle.setObjectName("pageSubtitle")
        header_layout.addWidget(subtitle)
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
            result = process_bank.match_payment_bills(
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

            process_bank.process_bank_statement(
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
                result = auto_entry.run_entry(
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
        process_bank.process_bank_statement(src, output_path=out, auto_summary=True)
        return 0

    sys.excepthook = _excepthook
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main() or 0)
