#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel报表工具 - PyQt5 图形界面
黑酷风格
"""

import sys
import os
from pathlib import Path
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QTextEdit, QProgressBar, QLineEdit,
    QFileDialog, QMessageBox, QGroupBox, QFrame, QComboBox
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QSize
from PyQt5.QtGui import QFont, QIcon, QPalette, QColor



# ═══════════════════════════════════════════════════════════════════════════════════
# 黑色主题样式
# ═══════════════════════════════════════════════════════════════════════════════════

DARK_STYLE = """
QMainWindow {
    background-color: #1a1a1a;
}
QWidget {
    background-color: #1a1a1a;
    color: #e0e0e0;
    font-family: "Microsoft YaHei UI", "Segoe UI", sans-serif;
    font-size: 13px;
}
QLabel {
    color: #e0e0e0;
    background-color: transparent;
}
QLabel#title {
    font-size: 22px;
    font-weight: bold;
    color: #ffffff;
    padding: 10px 0;
}
QLabel#subtitle {
    font-size: 12px;
    color: #888888;
}
QLineEdit {
    background-color: #2d2d2d;
    color: #ffffff;
    border: 1px solid #444444;
    border-radius: 4px;
    padding: 8px 12px;
    min-height: 28px;
}
QLineEdit:focus {
    border: 1px solid #00aaff;
}
QPushButton {
    background-color: #2d2d2d;
    color: #ffffff;
    border: 1px solid #444444;
    border-radius: 6px;
    padding: 10px 20px;
    min-width: 100px;
    font-weight: bold;
}
QPushButton:hover {
    background-color: #3d3d3d;
    border: 1px solid #00aaff;
}
QPushButton:pressed {
    background-color: #1d1d1d;
}
QPushButton#startBtn {
    background-color: #0078d4;
    border: 1px solid #0078d4;
}
QPushButton#startBtn:hover {
    background-color: #1a8ad4;
    border: 1px solid #1a8ad4;
}
QPushButton#startBtn:disabled {
    background-color: #1d3a54;
    color: #5588aa;
    border: 1px solid #1d3a54;
}
QPushButton#startBtn:pressed {
    background-color: #005fa3;
}
QPushButton#selectBtn {
    background-color: #2d2d2d;
    border: 1px solid #555555;
}
QPushButton#selectBtn:hover {
    border: 1px solid #00aaff;
}
QTextEdit {
    background-color: #0d0d0d;
    color: #b0b0b0;
    border: 1px solid #333333;
    border-radius: 6px;
    padding: 10px;
    font-family: "Consolas", "Microsoft YaHei UI", monospace;
    font-size: 12px;
}
QProgressBar {
    background-color: #2d2d2d;
    border: none;
    border-radius: 4px;
    height: 8px;
    text-align: center;
}
QProgressBar::chunk {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
        stop:0: #00aaff, stop:1: #00ffaa);
    border-radius: 4px;
}
QGroupBox {
    background-color: #1f1f1f;
    border: 1px solid #333333;
    border-radius: 8px;
    padding: 15px;
    margin-top: 10px;
}
QGroupBox::title {
    color: #aaaaaa;
    subcontrol-origin: margin;
    subcontrol-position: top left;
    padding: 0 8px;
}
QFrame#divider {
    background-color: #333333;
    max-height: 1px;
}
"""


# ═══════════════════════════════════════════════════════════════════════════════════
# 处理线程（后台运行，不阻塞UI）
# ═══════════════════════════════════════════════════════════════════════════════════

class ProcessThread(QThread):
    # 信号：状态更新、完成
    progress = pyqtSignal(str, int, int)   # text, current, total
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

        def callback(text, current, total):
            self.progress.emit(text, current, total)

        from report_maker import run_pipeline
        success, fail, results = run_pipeline(
            self.source_folder,
            self.n_batches,
            self.days_list,
            progress_callback=callback,
            cancel_check=self.is_cancel_requested
        )
        self.finished.emit(success, fail, results)



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
        self.setWindowTitle("Excel报表工具 v3.8")
        self.setMinimumSize(700, 600)
        self.thread = None

        self._setup_ui()
        self._load_default_path()

    def _setup_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(30, 20, 30, 20)
        main_layout.setSpacing(12)

        # 标题
        title = QLabel("Excel报表自动处理工具")
        title.setObjectName("title")
        title.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title)

        sub = QLabel("自动化六步流程 · 一键生成PDF报表")
        sub.setObjectName("subtitle")
        sub.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(sub)

        main_layout.addSpacing(5)

        # 分隔线
        div = QFrame()
        div.setObjectName("divider")
        div.setFixedHeight(1)
        main_layout.addWidget(div)

        main_layout.addSpacing(15)

        # === 文件夹选择区 ===
        folder_group = QGroupBox("📁 工程文件夹")
        folder_layout = QHBoxLayout(folder_group)
        folder_layout.setContentsMargins(10, 15, 10, 10)

        self.folder_edit = QLineEdit()
        self.folder_edit.setPlaceholderText("点击右侧按钮选择工程文件夹...")
        self.folder_edit.setReadOnly(True)

        self.select_btn = QPushButton("选择文件夹")
        self.select_btn.setObjectName("selectBtn")
        self.select_btn.setFixedWidth(110)
        self.select_btn.clicked.connect(self.select_folder)

        folder_layout.addWidget(self.folder_edit)
        folder_layout.addWidget(self.select_btn)
        main_layout.addWidget(folder_group)

        # === 参数输入区 ===
        param_group = QGroupBox("⚙️ 处理参数")
        param_layout = QVBoxLayout(param_group)
        param_layout.setContentsMargins(10, 15, 10, 10)
        param_layout.setSpacing(10)

        row1 = QHBoxLayout()
        row1.addWidget(QLabel("总期数:"))
        self.batch_input = QLineEdit()
        self.batch_input.setPlaceholderText("如: 2")
        self.batch_input.setFixedWidth(100)
        row1.addWidget(self.batch_input)
        row1.addSpacing(20)
        row1.addWidget(QLabel("每期间隔天数:"))
        self.days_input = QLineEdit()
        self.days_input.setPlaceholderText("如: 1,2,3 (逗号分隔)")
        row1.addWidget(self.days_input)
        row1.addStretch()
        param_layout.addLayout(row1)

        hint = QLabel("💡 提示：间隔天数可输入多个，用逗号分隔，例如 1,2,3 表示第1期隔1天，第2期隔2天...")
        hint.setStyleSheet("color: #666666; font-size: 11px;")
        param_layout.addWidget(hint)

        # Excel 类型选择
        excel_row = QHBoxLayout()
        excel_row.addWidget(QLabel("Excel 类型:"))
        self.excel_type_combo = QComboBox()
        self.excel_type_combo.addItems(["Microsoft Office", "WPS Office"])
        self.excel_type_combo.setCurrentIndex(0)
        self.excel_type_combo.setFixedWidth(160)
        excel_note = QLabel("（Office 静默运行，WPS 窗口会闪一下）")
        excel_note.setStyleSheet("color: #888888; font-size: 11px;")
        excel_row.addWidget(self.excel_type_combo)
        excel_row.addWidget(excel_note)
        excel_row.addStretch()
        param_layout.addLayout(excel_row)

        main_layout.addWidget(param_group)

        # === 按钮区 ===
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()

        self.start_btn = QPushButton("▶ 开始处理")
        self.start_btn.setObjectName("startBtn")
        self.start_btn.setFixedWidth(150)
        self.start_btn.setFixedHeight(42)
        self.start_btn.clicked.connect(self.start_process)
        btn_layout.addWidget(self.start_btn)

        self.stop_btn = QPushButton("✕ 停止")
        self.stop_btn.setFixedWidth(100)
        self.stop_btn.setEnabled(False)
        self.stop_btn.clicked.connect(self.stop_process)
        btn_layout.addWidget(self.stop_btn)

        main_layout.addLayout(btn_layout)

        # === 进度条 ===
        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setFormat("%p%")
        main_layout.addWidget(self.progress_bar)

        # === 日志区 ===
        log_label = QLabel("📋 处理日志")
        main_layout.addWidget(log_label)

        self.log_edit = QTextEdit()
        self.log_edit.setReadOnly(True)
        self.log_edit.setMaximumHeight(220)
        main_layout.addWidget(self.log_edit, 1)

        main_layout.addSpacing(10)

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

    def log(self, msg):
        self.log_edit.append(msg)
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
        self.log(f"🚀 开始处理: {n_batches}期，起始目录: {folder}")
        self.log(f"📊 间隔天数: {days_list}")

        self.start_btn.setEnabled(False)
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
        self.thread.finished.connect(self.on_finished)
        self.thread.start()

    def on_progress(self, text, current, total):
        self.log(f"  {text}")
        self.progress_bar.setValue(current)

    def on_finished(self, success, fail, results):
        self.progress_bar.setValue(self.progress_bar.maximum())

        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.select_btn.setEnabled(True)
        self.batch_input.setEnabled(True)
        self.days_input.setEnabled(True)
        self.excel_type_combo.setEnabled(True)

        cancelled = any(r.get("status") == "cancelled" for r in results)

        if cancelled:
            self.log("\n⚠️ 已安全停止，Excel资源已清理")
            QMessageBox.information(self, "已停止", "处理已在安全点停止。")
        elif fail == 0:
            self.log(f"\n🎉 全部完成！{success}期处理成功")
            QMessageBox.information(self, "完成", f"🎉 {success}期处理全部成功！")
        else:
            self.log(f"\n⚠️ 完成: {success}成功，{fail}失败")
            QMessageBox.warning(self, "部分失败", f"处理完成\n✅ {success}期成功\n❌ {fail}期失败\n\n请查看日志了解详情。")

        for r in results:
            if r["status"] == "ok":
                self.log(f"  ✅ {r['new_date']} 第{r['new_issue']}期 → PDF:{r['pdf_ok']}个")

        self.thread = None

    def stop_process(self):
        if self.thread and self.thread.isRunning():
            self.thread.request_stop()
            self.log("\n⚠️ 已请求停止，正在等待当前步骤安全收尾...")
        else:
            self.log("\n⚠️ 当前没有正在运行的任务")
        self.stop_btn.setEnabled(False)

    def closeEvent(self, event):
        if self.thread and self.thread.isRunning():
            self.thread.request_stop()
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
