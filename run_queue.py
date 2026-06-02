#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
多工程批量调度脚本
扫描 queue 文件夹，自动逐个处理其中的工程
每个工程文件夹下需有 config.txt 配置文件
日志输出到文件（同时实时输出到控制台）

选项：
  --yes                 自动确认，不等待用户输入
  --notify              处理完成后发 Windows Toast 通知
  --continue-on-error   某个工程失败时继续处理下一个
  --retry-failed        重新处理之前失败的工程（配合 --yes 使用）
  --dry-run             只显示要做什么，不实际执行
"""

import os
import re
import shutil
import subprocess
import sys
import logging
from pathlib import Path
from datetime import datetime

# 共享工具：颜色输出、IDE环境检测
from utils import green, red, yellow, cyan, bold

# ═══════════════════════════════════════════════════════════════════════════════════
# 路径配置（自动根据当前脚本位置推导，无需手动修改）
# ═══════════════════════════════════════════════════════════════════════════════════
PROJECT_DIR  = Path(__file__).resolve().parent         # 项目根目录
QUEUE_DIR    = PROJECT_DIR / "queue"                   # 待处理队列目录
DONE_DIR     = QUEUE_DIR / "done"                      # 已完成目录
REPORT_MAKER = PROJECT_DIR / "report_maker.py"         # 主程序路径
PYTHON       = sys.executable                          # 使用当前运行run_queue.py的Python
# ═══════════════════════════════════════════════════════════════════════════════════
# 内部配置
# ═══════════════════════════════════════════════════════════════════════════════════
LOG_DIR    = PROJECT_DIR / ".logs"
FAILED_LOG = LOG_DIR / "failed_projects.txt"


def setup_logging(project_name):
    """为每个工程设置独立的日志文件（tee效果：同时写文件和打印到控制台）"""
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime('%Y%m%d_%Hh%Mm%Ss')
    log_file = LOG_DIR / f"{project_name}_{timestamp}.log"

    # 创建 logger
    logger = logging.getLogger(project_name)
    logger.setLevel(logging.DEBUG)
    logger.handlers.clear()

    # 文件 handler
    file_handler = logging.FileHandler(log_file, encoding='utf-8')
    file_handler.setLevel(logging.DEBUG)

    # 控制台 handler（实时输出）
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)

    formatter = logging.Formatter('%(message)s')
    file_handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    return logger, log_file


def send_windows_notification(title, message):
    """发送 Windows Toast 通知"""
    try:
        from win10toast import ToastNotifier
        toaster = ToastNotifier()
        toaster.show_toast(title, message, duration=5, threaded=False)
    except ImportError:
        # 没有 win10toast，尝试用 PowerShell
        try:
            script = f'''
            [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
            [Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime] | Out-Null
            $template = @"
            <toast>
                <visual>
                    <binding template="ToastText02">
                        <text id="1">{title}</text>
                        <text id="2">{message}</text>
                    </binding>
                </visual>
            </toast>
"@
            $xml = New-Object Windows.Data.Xml.Dom.XmlDocument
            $xml.LoadXml($template)
            $toast = [Windows.UI.Notifications.ToastNotification]::new($xml)
            [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier("report_maker").Show($toast)
            '''
            subprocess.run(["powershell", "-Command", script],
                          capture_output=True, creationflags=0x08000000)
        except Exception:
            pass  # 通知失败不影响主流程


def load_failed_projects():
    """读取失败工程记录"""
    if not FAILED_LOG.exists():
        return set()
    return set(FAILED_LOG.read_text(encoding='utf-8').strip().splitlines())


def save_failed_projects(failed_set):
    """保存失败工程记录"""
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    FAILED_LOG.write_text('\n'.join(sorted(failed_set)), encoding='utf-8')


def parse_folder_info(name):
    """解析文件夹名称，提取期数和日期"""
    m = re.search(r'第(\d+)期\s+(\d{2})\.(\d{2})', name)
    return (int(m.group(1)), int(m.group(2)), int(m.group(3))) if m else (None, None, None)


def find_latest_issue(project_dir):
    """在工程文件夹里找最新一期"""
    subfolders = [d for d in project_dir.iterdir() if d.is_dir()]
    issue_folders = []
    for d in subfolders:
        info = parse_folder_info(d.name)
        if info[0] is not None:
            issue_folders.append((info[0], d))
    if not issue_folders:
        return None, None
    issue_folders.sort(key=lambda x: x[0])
    return issue_folders[-1]


def read_config(project_dir):
    """读取工程配置文件"""
    config_path = project_dir / "config.txt"
    if not config_path.exists():
        return None
    content = config_path.read_text(encoding='utf-8').strip()
    return content if content else None


def run_report_maker(project_dir, config_input, logger, dry_run=False):
    """运行 report_maker.py，传入配置参数"""
    if dry_run:
        logger.info(f"[模拟] 启动 report_maker.py（dry-run）...")
        logger.info(f"[模拟] 工程目录: {project_dir}")
        logger.info(f"[模拟] 参数: {config_input}")
        return 0

    logger.info(f"[调度] 启动 report_maker.py...")
    logger.info(f"[调度] 工程目录: {project_dir}")
    logger.info(f"[调度] 参数: {config_input}")

    cmd = [PYTHON, str(REPORT_MAKER)]
    env = os.environ.copy()
    env['REPORT_MAKER_PARENT'] = str(project_dir)
    env['PYTHONIOENCODING'] = 'utf-8'
    env['PYTHONUTF8'] = '1'

    process = subprocess.Popen(
        cmd,
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        encoding='utf-8',
        errors='replace',
        bufsize=1,
        env=env
    )

    process.stdin.write(f"{config_input}\n")
    process.stdin.flush()
    process.stdin.close()

    for line in iter(process.stdout.readline, ''):
        line = line.rstrip()
        if line:
            logger.info(line)

    return process.wait()


def print_banner():
    print("=" * 56)
    print("     多工程批量调度器")
    print("=" * 56)


def main():
    args = sys.argv[1:]
    auto_yes = '--yes' in args or '-y' in args
    continue_on_error = '--continue-on-error' in args or '-c' in args
    do_notify = '--notify' in args or '-n' in args
    retry_failed = '--retry-failed' in args
    dry_run = '--dry-run' in args or '--dryrun' in args

    # 清理已处理的参数
    clean_args = ['--yes', '-y', '--continue-on-error', '-c',
                  '--notify', '-n', '--retry-failed', '--dry-run', '--dryrun']
    args = [a for a in args if a not in clean_args]

    if '--help' in args or '-h' in args:
        print_banner()
        print("""
用法: python run_queue.py [选项]

选项:
  --yes, -y              自动确认，不等待用户输入
  --notify, -n            处理完成后发 Windows Toast 通知
  --continue-on-error    某个工程失败时继续处理下一个
  --retry-failed         重新处理之前失败的工程（需配合 --yes）
  --dry-run              只显示要做什么，不实际执行
  --help, -h             显示帮助信息

示例:
  python run_queue.py                       # 交互模式
  python run_queue.py --yes                 # 自动确认
  python run_queue.py --yes -n -c           # 完全无人值守+通知
  python run_queue.py --retry-failed --yes  # 重试失败工程
  python run_queue.py --dry-run             # 预览要处理的工程
""")
        return

    print_banner()

    if dry_run:
        print("\n[模拟运行] 只显示要做什么，不会实际执行任何操作\n")

    if not QUEUE_DIR.exists():
        print(f"[错误] 队列文件夹不存在: {QUEUE_DIR}")
        return

    if not REPORT_MAKER.exists():
        print(f"[错误] report_maker.py 不存在: {REPORT_MAKER}")
        return

    # 1. 找到所有待处理工程
    all_items = list(QUEUE_DIR.iterdir())
    project_folders = [d for d in all_items
                      if d.is_dir() and d.name not in ("done",)]
    skipped = [d for d in all_items if d.is_file()]

    if skipped:
        print(f"[警告] 跳过非文件夹: {[d.name for d in skipped]}")

    if not project_folders:
        print("[完成] 没有待处理的工程")
        return

    # --retry-failed：从 failed_projects.txt 读取历史失败工程
    if retry_failed:
        failed_set = load_failed_projects()
        if not failed_set:
            print("[完成] 没有失败工程记录（failed_projects.txt 为空或不存在）")
            return
        # 只处理记录中的工程（如果还在队列里）
        retry_projects = [p for p in project_folders if p.name in failed_set]
        if not retry_projects:
            print(f"[完成] 没有找到失败工程（可能已处理或已移除）")
            return
        print(f"[重试] 从失败记录中找到 {len(retry_projects)} 个工程:")
        for p in retry_projects:
            latest_issue, latest_folder = find_latest_issue(p)
            config = read_config(p)
            print(f"  - {p.name}  最新:{latest_folder.name if latest_folder else '?'}  参数:{config}")
        project_folders = retry_projects
        # 重试时强制自动确认
        auto_yes = True
        continue_on_error = True
    else:
        print(f"\n找到 {len(project_folders)} 个待处理工程:")
        for p in project_folders:
            latest_issue, latest_folder = find_latest_issue(p)
            config = read_config(p)
            config_display = config if config else "(无config.txt，跳过)"
            if latest_folder:
                print(f"  - {p.name}  最新:{latest_folder.name}  参数:{config_display}")
            else:
                print(f"  - {p.name}  {config_display}")

    # 检查有没有缺失 config.txt 的
    missing_config = [p for p in project_folders if not read_config(p)]
    if missing_config:
        print(f"\n[警告] 以下工程缺少 config.txt，将被跳过:")
        for p in missing_config:
            print(f"  - {p.name}")
        project_folders = [p for p in project_folders if read_config(p)]

    if not project_folders:
        print("[完成] 没有可处理的工程")
        return

    # 显示选项信息
    print("\n" + "-" * 56)
    print(f"[模式] {'自动确认' if auto_yes else '手动确认'} | {'失败继续' if continue_on_error else '失败中断'} | {'发通知' if do_notify else '不发通知'}")
    if dry_run:
        print("[模拟运行] 不会实际执行任何操作")
    print("-" * 56)

    # 确认是否开始
    if not auto_yes:
        print("\n是否开始按顺序处理？(y/n): ", end='', flush=True)
        confirm = input().strip().lower()
        if confirm != 'y':
            print("取消")
            return
    else:
        print("\n[自动] 开始处理...")

    # 统计
    success_count = 0
    fail_count = 0
    skip_count = 0
    failed_set = load_failed_projects()

    # 逐个处理
    for i, project in enumerate(project_folders, 1):
        latest_issue, latest_folder = find_latest_issue(project)
        config = read_config(project)

        # 设置日志（tee效果：同时写文件和输出到控制台）
        logger, log_file = setup_logging(project.name)
        logger.info("=" * 56)
        logger.info(f"  [{i}/{len(project_folders)}] 开始处理: {project.name}")
        logger.info(f"  最新一期: {latest_folder.name if latest_folder else '?'}")
        logger.info(f"  参数: {config}")
        logger.info("=" * 56)

        ret = run_report_maker(project, config, logger, dry_run=dry_run)

        if dry_run:
            # dry-run 不记录成功/失败
            logger.info(f"[模拟完成] {project.name}（模拟模式，不实际处理）")
            print(f"\n[模拟] {project.name} 完成（日志: {log_file}）")
        elif ret == 0:
            success_count += 1
            # 成功后从失败记录移除
            failed_set.discard(project.name)
            print(f"\n[完成] {project.name}  日志: {log_file}")
            logger.info(f"[完成] {project.name} 处理成功")
        else:
            # 记录失败
            failed_set.add(project.name)
            logger.warning(f"[警告] report_maker.py 异常退出 (code={ret})")
            if continue_on_error:
                fail_count += 1
                print(f"\n[失败] {project.name} (code={ret})，继续下一个...")
                logger.warning(f"[失败] {project.name} 处理失败，但继续执行下一个")
            else:
                fail_count += 1
                print(f"\n[中断] {project.name} 失败 (code={ret})，停止调度")
                logger.error(f"[中断] {project.name} 处理失败，调度中断")
                print(f"\n已完成: {success_count} 个 | 失败: {fail_count} 个 | 跳过: {skip_count} 个")
                print(f"失败工程: {project.name}")
                print(f"日志在: {log_file}")
                break

        # 处理完移动到已完成（dry-run 不移动）
        if not dry_run and not (continue_on_error and ret != 0):
            if not DONE_DIR.exists():
                DONE_DIR.mkdir(parents=True)

            done_path = DONE_DIR / project.name
            if done_path.exists():
                logger.warning(f"[警告] 已完成文件夹已存在: {done_path.name}，跳过移动")
                skip_count += 1
            else:
                try:
                    shutil.move(str(project), str(done_path))
                    logger.info(f"[完成] {project.name} 已移至 done/{done_path.name}")
                except Exception as e:
                    logger.warning(f"[警告] 移动失败: {e}")
                    skip_count += 1

    # 保存失败记录
    save_failed_projects(failed_set)

    # 汇总
    print("\n" + "=" * 56)
    print("  全部工程处理完毕！")
    print("=" * 56)
    print(f"成功: {success_count} 个")
    print(f"失败: {fail_count} 个")
    print(f"跳过: {skip_count} 个")
    if failed_set:
        print(f"\n失败工程记录在: {FAILED_LOG}")
        print(f"重试命令: python run_queue.py --retry-failed --yes")
    print(f"\n已完成的工程在: {DONE_DIR}")
    print(f"所有日志在: {LOG_DIR}")

    # 发通知
    if do_notify and not dry_run:
        title = "工程调度完成"
        if fail_count > 0:
            message = f"成功 {success_count} 个，失败 {fail_count} 个"
        else:
            message = f"全部 {success_count} 个工程处理完成！"
        send_windows_notification(title, message)


if __name__ == "__main__":
    main()
