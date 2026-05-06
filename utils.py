# -*- coding: utf-8 -*-
"""
共享工具模块
包含颜色输出、IDE环境检测等共享功能，供 report_maker.py 和 run_queue.py 共用
"""

import sys
import os

# ============ IDE 环境检测 ============
PYCHARM_MARKER = 'PyCharm' in sys.executable or 'pycharm' in sys.executable.lower()
IDE_ENV = (
    os.environ.get('PYCHARM_HOSTED') is not None or
    os.environ.get('PYCHARM') is not None or
    os.environ.get('VSCODE_PID') is not None or
    os.environ.get('VSCODE_INJECTION') is not None or
    PYCHARM_MARKER
)

# ============ 颜色输出 ============
# IDE 环境强制禁用颜色（PyCharm/VSCode 等不识别 ANSI 码）
# 非 IDE 环境：只在 TTY 终端下启用颜色
HAS_COLOR = sys.stdout is not None and sys.stdout.isatty() and not IDE_ENV

if HAS_COLOR:
    try:
        import colorama
        colorama.init(autoreset=True)
    except ImportError:
        HAS_COLOR = False


def _c(text, code):
    """给文本加 ANSI 颜色码"""
    return f"\033[{code}m{text}\033[0m" if HAS_COLOR else text


def green(text):   return _c(text, "92")
def red(text):     return _c(text, "91")
def yellow(text):  return _c(text, "93")
def cyan(text):    return _c(text, "96")
def bold(text):    return _c(text, "1")
