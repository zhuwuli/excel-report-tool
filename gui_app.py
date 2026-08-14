#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel报表工具 - PyQt5 图形界面
黑酷风格
"""

import sys
import os
import traceback
from html import escape
from pathlib import Path
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QGridLayout, QLabel, QPushButton, QTextEdit, QProgressBar, QLineEdit,
    QFileDialog, QMessageBox, QGroupBox, QFrame, QComboBox
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QSize
from PyQt5.QtGui import QFont, QIcon, QPalette, QColor


class SafeNullStream:
    """PyInstaller --windowed 下 stdout/stderr 可能为 None，先提供安全兜底。"""

    encoding = "utf-8"
    errors = "replace"

    def write(self, text):
        return 0

    def flush(self):
        pass

    def isatty(self):
        return False

    def writable(self):
        return True


if sys.stdout is None:
    sys.stdout = SafeNullStream()
if sys.stderr is None:
    sys.stderr = SafeNullStream()


APP_VERSION = "v3.12.0"
APP_TITLE = f"Excel报表工具 {APP_VERSION}"


# ═══════════════════════════════════════════════════════════════════════════════════
# 黑色主题样式
# ═══════════════════════════════════════════════════════════════════════════════════

DARK_STYLE = """
QMainWindow {
    background-color: #0f1419;
}
QWidget {
    background-color: #0f1419;
    color: #edf3f7;
    font-family: "Microsoft YaHei UI", "Segoe UI", sans-serif;
    font-size: 13px;
}
QFrame#heroPanel {
    background-color: #17212a;
    border: 1px solid #29485a;
    border-radius: 8px;
}
QFrame#sidePanel, QFrame#contentPanel {
    background-color: #171d23;
    border: 1px solid #28343e;
    border-radius: 8px;
}
QFrame#progressPanel, QFrame#metricCard, QFrame#stepCard, QFrame#logPanel {
    background-color: #111820;
    border: 1px solid #26333e;
    border-radius: 8px;
}
QFrame#logoBox {
    background-color: #0d3342;
    border: 1px solid #2696bd;
    border-radius: 8px;
}
QLabel {
    color: #edf3f7;
    background-color: transparent;
}
QLabel#logoText {
    color: #8ee7ff;
    font-size: 22px;
    font-weight: bold;
}
QLabel#title {
    font-size: 24px;
    font-weight: bold;
    color: #ffffff;
    padding: 0;
}
QLabel#subtitle {
    font-size: 12px;
    color: #9aaab6;
}
QLabel#versionBadge {
    color: #8ee7ff;
    background-color: #0d2a37;
    border: 1px solid #267ea4;
    border-radius: 11px;
    padding: 4px 12px;
    font-size: 11px;
    font-weight: bold;
}
QLabel#statusLabel {
    color: #c6f6ff;
    background-color: #102631;
    border: 1px solid #2a6a7f;
    border-radius: 6px;
    padding: 7px 12px;
    min-width: 170px;
}
QLabel#sectionTitle {
    color: #ffffff;
    font-size: 15px;
    font-weight: bold;
}
QLabel#fieldLabel, QLabel#metricTitle, QLabel#stepNumber {
    color: #9caab5;
    font-size: 12px;
}
QLabel#mutedText {
    color: #778792;
    font-size: 11px;
}
QLabel#summaryLabel {
    color: #dfe9ef;
    background-color: #101820;
    border: 1px solid #2a3944;
    border-radius: 7px;
    padding: 9px 10px;
}
QLabel#metricValue {
    color: #ffffff;
    font-size: 16px;
    font-weight: bold;
}
QLabel#stepTitle {
    color: #edf3f7;
    font-size: 13px;
    font-weight: bold;
}
QLabel#stepNumber {
    background-color: #0d2a37;
    border: 1px solid #267ea4;
    border-radius: 5px;
    padding: 3px 0;
}
QLabel#statusDotReady {
    color: #33d6a6;
    font-size: 15px;
}
QLineEdit, QComboBox {
    background-color: #0c1218;
    color: #ffffff;
    border: 1px solid #344653;
    border-radius: 6px;
    padding: 7px 10px;
    min-height: 28px;
}
QLineEdit:focus, QComboBox:focus {
    border: 1px solid #31b7f0;
    background-color: #0d1820;
}
QLineEdit:disabled, QComboBox:disabled {
    color: #778590;
    background-color: #171f26;
}
QPushButton {
    background-color: #24313b;
    color: #ffffff;
    border: 1px solid #3a4c59;
    border-radius: 7px;
    padding: 9px 16px;
    font-weight: bold;
}
QPushButton:hover {
    background-color: #2d3b46;
    border: 1px solid #31b7f0;
}
QPushButton:pressed {
    background-color: #1c2933;
}
QPushButton#startBtn {
    background-color: #0787d8;
    border: 1px solid #28aaf0;
    font-size: 15px;
}
QPushButton#startBtn:hover {
    background-color: #139ce8;
}
QPushButton#startBtn:disabled {
    background-color: #1c3a52;
    color: #7292a8;
    border: 1px solid #2c4b60;
}
QPushButton#selectBtn {
    background-color: #1f2b34;
}
QPushButton#stopBtn {
    background-color: #3b2528;
    border: 1px solid #834349;
}
QPushButton#stopBtn:hover {
    background-color: #4b2b30;
    border: 1px solid #e06f78;
}
QPushButton#stopBtn:disabled {
    background-color: #20272d;
    color: #737f88;
    border: 1px solid #303c46;
}
QTextEdit {
    background-color: #090e13;
    color: #b8c6cf;
    border: 1px solid #273541;
    border-radius: 8px;
    padding: 12px;
    font-family: "Consolas", "Microsoft YaHei UI", monospace;
    font-size: 12px;
}
QProgressBar {
    background-color: #0b1117;
    border: 1px solid #2a3944;
    border-radius: 6px;
    min-height: 16px;
    text-align: center;
    color: #e4eef4;
}
QProgressBar::chunk {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
        stop:0: #1e9bf0, stop:1: #20d6b0);
    border-radius: 5px;
}
"""


# ═══════════════════════════════════════════════════════════════════════════════════
# 处理线程（后台运行，不阻塞UI）
# ═══════════════════════════════════════════════════════════════════════════════════

class ProcessThread(QThread):
    # 信号：状态更新、完成
    progress = pyqtSignal(str, int, int)   # text, current, total
    log_output = pyqtSignal(str)           # 后台print/stderr输出
    finished = pyqtSignal(int, int, list)  # success, fail, results

    def __init__(self, source_folder, n_batches, days_list, excel_type="office"):
        super().__init__()
        self.source_folder = source_folder
        self.n_batches = n_batches
        self.days_list = days_list
        self.excel_type = excel_type
        self._cancel_requested = False

    def request_stop(self):
        """请求后台流程在下一个安全点停止"""
        self._cancel_requested = True

    def is_cancel_requested(self):
        return self._cancel_requested

    def run(self):
        import warnings
        warnings.filterwarnings("ignore")
        os.environ["RPT_EXCEL_TYPE"] = self.excel_type
        old_stdout, old_stderr = sys.stdout, sys.stderr
        gui_stream = GuiLogStream(self.log_output.emit)

        def callback(text, current, total):
            self.progress.emit(text, current, total)

        try:
            sys.stdout = gui_stream
            sys.stderr = gui_stream

            from report_maker import run_pipeline
            success, fail, results = run_pipeline(
                self.source_folder,
                self.n_batches,
                self.days_list,
                progress_callback=callback,
                cancel_check=self.is_cancel_requested
            )
        except Exception as e:
            gui_stream.write("\n[严重错误] 后台任务异常退出:\n")
            gui_stream.write(traceback.format_exc())
            success, fail = 0, 1
            results = [{
                "status": "fail",
                "error": str(e),
            }]
        finally:
            gui_stream.flush()
            sys.stdout = old_stdout
            sys.stderr = old_stderr

        self.finished.emit(success, fail, results)


class GuiLogStream:
    """把后台print/stderr按行转发到GUI日志框。"""

    encoding = "utf-8"
    errors = "replace"

    def __init__(self, emit_func):
        self.emit_func = emit_func
        self._buffer = ""

    def isatty(self):
        return False

    def writable(self):
        return True

    def write(self, text):
        if not text:
            return
        text = str(text).replace("\r", "\n")
        self._buffer += text
        while "\n" in self._buffer:
            line, self._buffer = self._buffer.split("\n", 1)
            if line.strip():
                self.emit_func(line)

    def flush(self):
        if self._buffer.strip():
            self.emit_func(self._buffer.rstrip())
        self._buffer = ""


# ═══════════════════════════════════════════════════════════════════════════════════
# Windows 原生文件夹选择对话框（避免 QFileDialog 无响应）
# ═══════════════════════════════════════════════════════════════════════════════════




# ═══════════════════════════════════════════════════════════════════════════════════
# Windows 原生文件夹选择对话框（避免 Qt/Win32 COM 冲突）
# ═══════════════════════════════════════════════════════════════════════════════════

# ═══════════════════════════════════════════════════════════════════════════════════
# 主窗口
# ═══════════════════════════════════════════════════════════════════════════════════

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.setMinimumSize(1060, 720)
        self.resize(1180, 780)
        self.thread = None

        self._setup_ui()
        self._load_default_path()
        self._update_task_summary()

    def _setup_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(26, 24, 26, 24)
        root.setSpacing(18)

        # === 顶部品牌区 ===
        hero = QFrame()
        hero.setObjectName("heroPanel")
        hero_layout = QHBoxLayout(hero)
        hero_layout.setContentsMargins(22, 16, 22, 16)
        hero_layout.setSpacing(18)

        logo_box = QFrame()
        logo_box.setObjectName("logoBox")
        logo_box.setFixedSize(48, 48)
        logo_layout = QVBoxLayout(logo_box)
        logo_layout.setContentsMargins(0, 0, 0, 0)
        logo = QLabel("XL")
        logo.setObjectName("logoText")
        logo.setAlignment(Qt.AlignCenter)
        logo_layout.addWidget(logo)
        hero_layout.addWidget(logo_box)

        title_block = QVBoxLayout()
        title_block.setSpacing(4)
        title = QLabel("Excel 报表工作台")
        title.setObjectName("title")
        subtitle = QLabel("把工程复制、内容更新、外链刷新和 PDF 导出收进一个稳定流程")
        subtitle.setObjectName("subtitle")
        title_block.addWidget(title)
        title_block.addWidget(subtitle)
        hero_layout.addLayout(title_block, 1)

        right_head = QVBoxLayout()
        right_head.setSpacing(8)
        right_head.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        version = QLabel(APP_VERSION)
        version.setObjectName("versionBadge")
        version.setAlignment(Qt.AlignCenter)
        self.status_label = QLabel("就绪")
        self.status_label.setObjectName("statusLabel")
        self.status_label.setAlignment(Qt.AlignCenter)
        right_head.addWidget(version, 0, Qt.AlignRight)
        right_head.addWidget(self.status_label, 0, Qt.AlignRight)
        hero_layout.addLayout(right_head)
        root.addWidget(hero)

        # === 主工作区 ===
        body = QHBoxLayout()
        body.setSpacing(16)
        root.addLayout(body, 1)

        # 左侧：任务配置
        side = QFrame()
        side.setObjectName("sidePanel")
        side.setFixedWidth(420)
        side_layout = QVBoxLayout(side)
        side_layout.setContentsMargins(18, 18, 18, 18)
        side_layout.setSpacing(9)

        config_title = QLabel("任务配置")
        config_title.setObjectName("sectionTitle")
        side_layout.addWidget(config_title)

        folder_label = QLabel("01  选择工程文件夹")
        folder_label.setObjectName("fieldLabel")
        side_layout.addWidget(folder_label)

        self.folder_edit = QLineEdit()
        self.folder_edit.setPlaceholderText("请选择起始工程文件夹")
        self.folder_edit.setMinimumHeight(34)
        self.folder_edit.setReadOnly(True)
        side_layout.addWidget(self.folder_edit)

        self.select_btn = QPushButton("选择文件夹")
        self.select_btn.setObjectName("selectBtn")
        self.select_btn.setMinimumHeight(34)
        self.select_btn.clicked.connect(self.select_folder)
        side_layout.addWidget(self.select_btn)

        batch_section = QLabel("02  设置处理批次")
        batch_section.setObjectName("fieldLabel")
        side_layout.addWidget(batch_section)

        param_grid = QGridLayout()
        param_grid.setHorizontalSpacing(10)
        param_grid.setVerticalSpacing(10)

        batch_label = QLabel("总期数")
        batch_label.setObjectName("fieldLabel")
        self.batch_input = QLineEdit()
        self.batch_input.setPlaceholderText("如 2")
        self.batch_input.setMinimumWidth(105)
        param_grid.addWidget(batch_label, 0, 0)
        param_grid.addWidget(self.batch_input, 1, 0)

        days_label = QLabel("每期间隔天数")
        days_label.setObjectName("fieldLabel")
        self.days_input = QLineEdit()
        self.days_input.setPlaceholderText("如 1 或 1,2,3")
        param_grid.addWidget(days_label, 0, 1)
        param_grid.addWidget(self.days_input, 1, 1)
        param_grid.setColumnStretch(0, 0)
        param_grid.setColumnStretch(1, 1)
        side_layout.addLayout(param_grid)

        excel_label = QLabel("03  选择 Excel 类型")
        excel_label.setObjectName("fieldLabel")
        side_layout.addWidget(excel_label)
        self.excel_type_combo = QComboBox()
        self.excel_type_combo.addItems(["Microsoft Office", "WPS Office"])
        self.excel_type_combo.setCurrentIndex(0)
        self.excel_type_combo.setMinimumHeight(34)
        side_layout.addWidget(self.excel_type_combo)

        note = QLabel("Office 静默运行；WPS 可能短暂显示窗口。")
        note.setObjectName("mutedText")
        note.setWordWrap(False)
        side_layout.addWidget(note)

        summary_title = QLabel("任务摘要")
        summary_title.setObjectName("fieldLabel")
        side_layout.addWidget(summary_title)

        self.summary_label = QLabel("等待填写任务参数")
        self.summary_label.setObjectName("summaryLabel")
        self.summary_label.setWordWrap(True)
        self.summary_label.setMinimumHeight(64)
        self.summary_label.setMaximumHeight(86)
        side_layout.addWidget(self.summary_label)

        action_title = QLabel("04  启动任务")
        action_title.setObjectName("fieldLabel")
        side_layout.addWidget(action_title)

        self.start_btn = QPushButton("▶ 开始处理")
        self.start_btn.setObjectName("startBtn")
        self.start_btn.setFixedHeight(42)
        self.start_btn.clicked.connect(self.start_process)
        side_layout.addWidget(self.start_btn)

        self.stop_btn = QPushButton("■ 停止")
        self.stop_btn.setObjectName("stopBtn")
        self.stop_btn.setFixedHeight(36)
        self.stop_btn.setEnabled(False)
        self.stop_btn.clicked.connect(self.stop_process)
        side_layout.addWidget(self.stop_btn)

        self.batch_input.textChanged.connect(self._update_task_summary)
        self.days_input.textChanged.connect(self._update_task_summary)
        self.excel_type_combo.currentIndexChanged.connect(self._update_task_summary)

        body.addWidget(side)

        # 右侧：运行状态和日志
        content = QFrame()
        content.setObjectName("contentPanel")
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(20, 20, 20, 20)
        content_layout.setSpacing(16)

        metrics = QHBoxLayout()
        metrics.setSpacing(12)
        self.current_metric = self._make_metric("当前步骤", "等待开始")
        self.mode_metric = self._make_metric("运行模式", "Office")
        self.result_metric = self._make_metric("处理结果", "暂无结果")
        metrics.addWidget(self.current_metric[0])
        metrics.addWidget(self.mode_metric[0])
        metrics.addWidget(self.result_metric[0])
        content_layout.addLayout(metrics)

        progress_frame = QFrame()
        progress_frame.setObjectName("progressPanel")
        progress_layout = QVBoxLayout(progress_frame)
        progress_layout.setContentsMargins(14, 12, 14, 12)
        progress_title = QLabel("处理进度")
        progress_title.setObjectName("sectionTitle")
        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setFormat("%p%")
        progress_layout.addWidget(progress_title)
        progress_layout.addWidget(self.progress_bar)
        content_layout.addWidget(progress_frame)

        log_panel = QFrame()
        log_panel.setObjectName("logPanel")
        log_panel_layout = QVBoxLayout(log_panel)
        log_panel_layout.setContentsMargins(14, 12, 14, 14)
        log_panel_layout.setSpacing(10)

        log_header = QHBoxLayout()
        log_title = QLabel("实时日志")
        log_title.setObjectName("sectionTitle")
        self.log_dot = QLabel("●")
        self.log_dot.setObjectName("statusDotReady")
        log_hint = QLabel("错误、警告和完成信息会在这里高亮")
        log_hint.setObjectName("mutedText")
        log_header.addWidget(self.log_dot)
        log_header.addWidget(log_title)
        log_header.addStretch()
        log_header.addWidget(log_hint)
        log_panel_layout.addLayout(log_header)

        self.log_edit = QTextEdit()
        self.log_edit.setReadOnly(True)
        log_panel_layout.addWidget(self.log_edit, 1)
        content_layout.addWidget(log_panel, 1)

        body.addWidget(content, 1)

    def _make_metric(self, title, value):
        card = QFrame()
        card.setObjectName("metricCard")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(16, 14, 16, 14)
        layout.setSpacing(6)
        title_label = QLabel(title)
        title_label.setObjectName("metricTitle")
        value_label = QLabel(value)
        value_label.setObjectName("metricValue")
        value_label.setWordWrap(True)
        layout.addWidget(title_label)
        layout.addWidget(value_label)
        return card, value_label

    def _make_step_card(self, number, title, desc):
        card = QFrame()
        card.setObjectName("stepCard")
        layout = QHBoxLayout(card)
        layout.setContentsMargins(12, 8, 12, 8)
        layout.setSpacing(10)

        number_label = QLabel(number)
        number_label.setObjectName("stepNumber")
        number_label.setFixedWidth(28)
        number_label.setAlignment(Qt.AlignCenter)

        text_layout = QVBoxLayout()
        text_layout.setSpacing(2)
        title_label = QLabel(title)
        title_label.setObjectName("stepTitle")
        desc_label = QLabel(desc)
        desc_label.setObjectName("mutedText")
        desc_label.setWordWrap(False)
        text_layout.addWidget(title_label)
        text_layout.addWidget(desc_label)

        layout.addWidget(number_label)
        layout.addLayout(text_layout, 1)
        return card

    def _load_default_path(self):
        """加载默认路径到界面"""
        try:
            from report_maker import BASE_DIR, find_latest_folder, parse_folder_info
            base = Path(BASE_DIR)
            if base.exists():
                latest_issue, latest_folder = find_latest_folder(base)
                if latest_folder:
                    self.folder_edit.setText(str(latest_folder))
                    issue, _, _ = parse_folder_info(latest_folder.name)
                    if issue:
                        self.log(f"✅ 检测到最新一期: {latest_folder.name}，可作为起始文件夹")
        except Exception:
            pass

    def _update_task_summary(self):
        folder = self.folder_edit.text().strip() or "未选择"
        batches = self.batch_input.text().strip() or "未填写"
        days = self.days_input.text().strip() or "未填写"
        excel_name = self.excel_type_combo.currentText() if hasattr(self, "excel_type_combo") else "Microsoft Office"
        mode_name = "WPS" if "WPS" in excel_name else "Office"
        self.summary_label.setText(
            f"目录：{folder}\n"
            f"期数：{batches}    间隔：{days}    模式：{excel_name}"
        )
        if hasattr(self, "mode_metric"):
            self.mode_metric[1].setText(mode_name)

    def _set_status(self, text, color="#c8f7ff"):
        self.status_label.setText(text)
        self.status_label.setStyleSheet(
            f"color: {color}; background-color: #122530; border: 1px solid #2d6174; "
            "border-radius: 6px; padding: 7px 12px; min-width: 170px;"
        )
        if hasattr(self, "log_dot"):
            self.log_dot.setStyleSheet(f"color: {color}; font-size: 15px;")

    def _log_color(self, msg):
        if any(key in msg for key in ("[严重错误]", "[失败]", "❌", "Traceback", "Error", "Exception")):
            return "#ff8f8f"
        if any(key in msg for key in ("[警告]", "[提醒]", "⚠️")):
            return "#ffd37a"
        if any(key in msg for key in ("[成功]", "[完成]", "✅", "🎉")):
            return "#8ff0a4"
        if any(key in msg for key in ("[处理]", "[工具]", "[文件]", "🚀", "📊", "📦")):
            return "#8fdcff"
        return "#b0b0b0"

    def log(self, msg):
        color = self._log_color(str(msg))
        html = escape(str(msg)).replace("\n", "<br>")
        self.log_edit.append(f'<span style="color:{color};">{html}</span>')
        # 自动滚动到底部
        scrollbar = self.log_edit.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(
            self, "选择工程文件夹", str(Path.home()),
            QFileDialog.ShowDirsOnly
        )
        if folder:
            self.folder_edit.setText(folder)
            self._update_task_summary()
            self.log(f"📂 已选择: {folder}")

    def start_process(self):
        from report_maker import run_pipeline, find_latest_folder, parse_folder_info
        folder = self.folder_edit.text().strip()
        if not folder or not Path(folder).exists():
            QMessageBox.warning(self, "警告", "请先选择有效的工程文件夹！")
            return

        batch_text = self.batch_input.text().strip()
        days_text = self.days_input.text().strip()

        if not batch_text:
            QMessageBox.warning(self, "警告", "请输入总期数！")
            return
        if not days_text:
            QMessageBox.warning(self, "警告", "请输入间隔天数！")
            return

        try:
            n_batches = int(batch_text)
            if n_batches <= 0:
                QMessageBox.warning(self, "警告", "总期数必须大于 0！")
                return
            days_parts = [x.strip() for x in days_text.split(",")]
            if len(days_parts) == 1:
                days_list = [int(days_parts[0])] * n_batches
            else:
                if len(days_parts) != n_batches:
                    QMessageBox.warning(self, "警告", f"间隔天数个数({len(days_parts)})需等于总期数({n_batches})！")
                    return
                days_list = [int(x) for x in days_parts]
        except ValueError:
            QMessageBox.warning(self, "警告", "输入格式错误！\n总期数填数字，间隔天数用逗号分隔。")
            return

        self.log_edit.clear()
        self._set_status("运行中 · 正在准备处理任务", "#8fdcff")
        self.current_metric[1].setText("准备中")
        self.result_metric[1].setText("运行中")
        self._update_task_summary()
        self.log(f"🚀 开始处理: {n_batches}期，起始目录: {folder}")
        self.log(f"📊 间隔天数: {days_list}")

        self.start_btn.setEnabled(False)
        self.stop_btn.setText("■ 停止")
        self.stop_btn.setEnabled(True)
        self.select_btn.setEnabled(False)
        self.batch_input.setEnabled(False)
        self.days_input.setEnabled(False)
        self.excel_type_combo.setEnabled(False)

        self.progress_bar.setValue(0)
        self.progress_bar.setMaximum(n_batches)

        excel_type = "wps" if self.excel_type_combo.currentIndex() == 1 else "office"
        self.log(f"📦 Excel 类型: {excel_type}")

        self.thread = ProcessThread(folder, n_batches, days_list, excel_type=excel_type)
        self.thread.progress.connect(self.on_progress)
        self.thread.log_output.connect(self.on_thread_log)
        self.thread.finished.connect(self.on_finished)
        self.thread.start()

    def on_progress(self, text, current, total):
        self.log(f"  {text}")
        self.progress_bar.setValue(current)
        self.current_metric[1].setText(f"{current}/{total}")
        self._set_status(f"运行中 · {text}", "#8fdcff")

    def on_thread_log(self, text):
        self.log(text)

    def on_finished(self, success, fail, results):
        self.progress_bar.setValue(self.progress_bar.maximum())

        self.start_btn.setEnabled(True)
        self.stop_btn.setText("■ 停止")
        self.stop_btn.setEnabled(False)
        self.select_btn.setEnabled(True)
        self.batch_input.setEnabled(True)
        self.days_input.setEnabled(True)
        self.excel_type_combo.setEnabled(True)

        cancelled = any(r.get("status") == "cancelled" for r in results)

        if cancelled:
            self._set_status("已停止 · Excel资源已清理", "#ffd37a")
            self.current_metric[1].setText("已停止")
            self.result_metric[1].setText(f"成功 {success} / 失败 {fail}")
            self.log("\n⚠️ 已安全停止，Excel资源已清理")
            QMessageBox.information(self, "已停止", "处理已在安全点停止。")
        elif fail == 0:
            self._set_status(f"完成 · {success}期全部成功", "#8ff0a4")
            self.current_metric[1].setText("已完成")
            self.result_metric[1].setText(f"成功 {success} 期")
            self.log(f"\n🎉 全部完成！{success}期处理成功")
            QMessageBox.information(self, "完成", f"🎉 {success}期处理全部成功！")
        else:
            self._set_status(f"部分失败 · {success}成功 / {fail}失败", "#ffd37a")
            self.current_metric[1].setText("已结束")
            self.result_metric[1].setText(f"成功 {success} / 失败 {fail}")
            self.log(f"\n⚠️ 完成: {success}成功，{fail}失败")
            QMessageBox.warning(self, "部分失败", f"处理完成\n✅ {success}期成功\n❌ {fail}期失败\n\n请查看日志了解详情。")

        for r in results:
            if r["status"] == "ok":
                self.log(f"  ✅ {r['new_date']} 第{r['new_issue']}期 → PDF:{r['pdf_ok']}个")

        self.thread = None

    def stop_process(self):
        if self.thread and self.thread.isRunning():
            self.thread.request_stop()
            self.stop_btn.setText("正在停止...")
            self.current_metric[1].setText("停止中")
            self._set_status("正在停止 · 等待当前步骤安全收尾", "#ffd37a")
            self.log("\n⚠️ 已请求停止，正在等待当前步骤安全收尾...")
        else:
            self.log("\n⚠️ 当前没有正在运行的任务")
        self.stop_btn.setEnabled(False)

    def closeEvent(self, event):
        if self.thread and self.thread.isRunning():
            self.thread.request_stop()
            self._set_status("正在停止 · 停止完成后可关闭窗口", "#ffd37a")
            self.log("\n⚠️ 正在处理任务，已请求安全停止；停止完成后再关闭窗口。")
            event.ignore()
            return
        event.accept()


# ═══════════════════════════════════════════════════════════════════════════════════
# 入口
# ═══════════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("fusion")
    app.setStyleSheet(DARK_STYLE)
    win = MainWindow()
    win.show()
    sys.exit(app.exec_())
