# 变更日志

所有版本变更记录，按时间倒序排列。

---

## v3.11.3 (2026-07-22) - WPS PDF export guard and #### column-width fix

### Background
- Tested on 5 PCs: 2 Office PCs were normal; 2 of 3 WPS PCs produced extra PDF pages during automated export.
- Manual WPS PDF export could still be normal, so the main risk is WPS COM `ExportAsFixedFormat()` pagination.
- Some cells could also appear as `####` in the exported PDF when their displayed text overflowed the current column width.

### PDF Export Guard
- Added `_fix_hash_cells_before_pdf()` before PDF export.
- Detects `####` by reading COM `cell.Text`, not by searching `cell.Value`.
- Does not change font size and does not use `ShrinkToFit`.
- Only widens affected columns, up to 3 attempts per column.
- Each widening step uses `max(old_width + 2, old_width * 1.15)`, capped at width 30.
- If any column width changes, the generated Excel workbook is saved with `wb.Save()`.

### WPS Pagination Guard
- Added `_prepare_wps_pages_for_pdf_export()`, enabled only in WPS mode.
- Office mode keeps the previous pagination behavior.
- WPS visible sheets use A4, `Zoom = False`, and `FitToPagesWide = 1`.
- Fixed one-page sheets use `FitToPagesTall = 1`; data sheets keep natural vertical pagination.
- Does not hard-code expected PDF page count, compare with previous PDFs, or globally clear page breaks.

### Logs and Builds
- Logs now include `found / resolved / unresolved` for `####` cells.
- Logs include sheet/column, width before/after, attempts, affected cell count, unresolved count, and save status.
- Morning test builds: `ExcelReportTool_v3.11.1_Windows_x64`, `ExcelReportTool_v3.11.2_Windows_x64`.
- Current recommended build: `ExcelReportTool_v3.11.3_Windows_x64`.

### Verification
- `python -m py_compile report_maker.py gui_app.py utils.py` passed.
- `_is_hash_display()` checks passed.
- Mock COM tests passed for visible-sheet filtering, hidden-sheet skipping, width cap 30, and WPS PageSetup behavior.
- Real Excel/WPS batch export was not run automatically to avoid closing active user workbooks.

---

## v3.11 (2026-07-03) - GUI 重设计 + 无黑框发布体验

### GUI 重设计
- `gui_app.py` 界面版本号更新为 `v3.11`
- 界面改为工作台布局：顶部品牌区、左侧任务配置、右侧运行状态与实时日志
- 增加运行状态卡片：当前步骤、运行模式、处理结果
- 实时日志支持颜色区分：成功、警告、失败、处理中信息更清晰
- 左侧布局经过多轮调整，修复小窗口下文字重叠、按钮压住内容的问题

### 无黑框与错误可见性
- 新增 `SafeNullStream`，适配 PyInstaller `--windowed` 下 `sys.stdout/sys.stderr` 可能为 `None` 的情况
- 新增 `GuiLogStream`，后台 `print()`、`stderr` 和异常 traceback 会转入 GUI 日志区
- `_kill_excel_process()` 调用 `taskkill` 时使用 `CREATE_NO_WINDOW` 和 `STARTUPINFO` 隐藏子进程窗口
- 推荐发布 exe 名称：`ExcelReportTool_v3.11.3_Windows_x64.exe`

### 推荐打包命令
```powershell
python -m PyInstaller --onefile --clean --windowed `
  --hidden-import win32timezone `
  --hidden-import win32crypt `
  --hidden-import pywinauto `
  --hidden-import pywinauto.keyboard `
  --hidden-import pywinauto.timings `
  --hidden-import pywinauto.application `
  --hidden-import pywinauto.findwindows `
  --name "ExcelReportTool_v3.11.3_Windows_x64" gui_app.py
```

### 验证
- `python -m py_compile gui_app.py report_maker.py utils.py` 通过
- PyInstaller `--windowed` 导入模拟通过
- `MainWindow` 实例化测试通过
- `git diff --check` 通过

---

## v3.10 Hotfix (2026-06-27) - STEP4/STEP5 性能优化

### STEP4：Range 批量迁移
- 日期列迁移由逐单元格 COM 读写改为整段 `Range.Value` 批量读写
- 保留“第3列为空时不覆盖第2列已有值”的原业务规则
- 大幅减少跨进程 COM 调用，数据行越多收益越明显

### STEP5：智能等待
- 启动固定等待 2 秒改为检测 `Ready`：至少 0.5 秒、最多 2 秒
- 使用 `LinkSources(1)` 检测外部链接，无链接时跳过按钮等待
- 修复 `wait_until_passes()` 不会按 `False` 返回值持续等待的问题
- 更新按钮每 0.25 秒轮询，Office/WPS 上限分别为 6 秒/8 秒
- 更新后检测 `CalculationState`：至少 0.5 秒、最多 5 秒
- 文件间隔由 1 秒缩短为 0.2 秒，新增单文件耗时日志
- 保留逐文件实例、保存关闭、进程清理、重试和安全停止逻辑

### 验证与备份
- 语法检查及 STEP4/STEP5 模拟测试通过
- 未自动运行真实 Excel COM 测试，避免关闭用户当前工作簿
- STEP4 备份：`.backup\20260626_164422\report_maker.before-step4-range.py`
- STEP5 备份：`.backup\20260627_090812\report_maker.before-step5-smart-wait.py`

---

## v3.10 Hotfix (2026-06-26) - 开发环境与打包修复

### 环境与原因
- 旧 `.venv` 指向已不存在的 `C:\Python314`，使用 `D:\mysoftware\python\python.exe`（Python 3.8.6）重建
- PyQt5、openpyxl、win32com、pywinauto、requests、PyInstaller、tqdm 导入验证通过
- `(.venv) (base)` 共存时，直接执行 `pyinstaller` 实际可能命中 `F:\anaconda\Scripts\pyinstaller.exe`
- `--collect-all pywinauto` 配合复杂 Anaconda 环境可能收集 pandas、scipy、sklearn、Jupyter 等无关包，导致 exe 增长到约 600 MB

### 当前推荐打包命令
```powershell
python -m PyInstaller --onefile --clean `
  --hidden-import win32timezone `
  --hidden-import win32crypt `
  --hidden-import pywinauto `
  --hidden-import pywinauto.keyboard `
  --hidden-import pywinauto.timings `
  --hidden-import pywinauto.application `
  --hidden-import pywinauto.findwindows `
  --name "Excel报表工具_v3.10_hotfix" gui_app.py
```

---

## v3.10 Hotfix (2026-06-26) - 鲁棒性增强

### 问题修复
- **WPS 安全停止后可能仍残留表格进程**
  - 原因：原清理逻辑只固定清理 `EXCEL.EXE`，WPS 表格常见进程为 `et.exe`
  - 解决：新增 `_excel_process_names()`，WPS 模式下同时尝试清理 `EXCEL.EXE` 和 `et.exe`
- **`~$` 临时锁文件会被复制到新一期目录**
  - 原因：旧逻辑只在查找 Excel 文件时跳过 `~$`，但 `shutil.copytree()` 会原样复制隐藏临时文件
  - 解决：复制目录时通过 `ignore` 回调过滤 `~$*.xls` / `~$*.xlsx` / `~$*.xlsm`

### 新增功能
- **运行前轻量自检**
  - 正式处理前检查工程路径、最新一期文件夹、汇总表、Excel 类型和自动化依赖
  - 如果最新一期存在 `~$` 临时锁文件，会提示但不中断；程序会跳过并避免复制到新目录
  - GUI 和命令行入口都会先执行自检

### 代码修改
- `report_maker.py`：新增 `TEMP_LOCK_PREFIX`、`TEMP_EXCEL_EXTS` 临时文件配置
- `report_maker.py`：新增 `is_temp_excel_file()`、`find_temp_excel_files()`、`_ignore_temp_excel_files()`
- `report_maker.py`：`copy_directory()` 复制时过滤临时锁文件
- `report_maker.py`：`_kill_excel_process()` 支持按 Excel 类型清理 Office/WPS 进程
- `report_maker.py`：新增 `preflight_check()`、`print_preflight_result()`，并接入 `run_pipeline()` 和 `main()`

### 验证
- `python -m py_compile report_maker.py gui_app.py run_queue.py utils.py` 通过
- 验证 WPS 模式进程名单包含 `et.exe`
- 验证复制新目录时 `~$` 临时锁文件不会被带入
- 验证运行前自检可识别工程父目录和某一期目录

---

## v3.10 Hotfix (2026-06-02) - 临时文件过滤 + 安全停止 + 路径自动化

### 问题修复
- **WPS提示无法打开 `~$汇总表.xlsx`**
  - 原因：`find_excel_files()` 只按扩展名和“汇总”关键字筛选，未排除 WPS/Excel 生成的 `~$` 临时锁文件
  - 现象：目录里残留 `~$汇总表.xlsx` 时，程序会误把它当作正式汇总表打开，WPS 弹出“无法打开文件”
  - 解决：`find_excel_files()` 增加 `f.startswith('~$')` 过滤，自动跳过临时锁文件
- **GUI停止按钮容易残留Excel进程**
  - 原因：原逻辑使用 `QThread.terminate()` 硬终止后台线程，Excel COM 清理代码可能来不及执行
  - 解决：改为协作式安全停止，GUI 设置取消标志，`report_maker.py` 在步骤边界检查并安全退出
  - 清理：`process_excel()` 增加 `finally`，取消或异常时尽量执行 `Workbook.Close(SaveChanges=False)`、`Excel.Quit()` 和 `_kill_excel_process()`

### 代码修改
- `gui_app.py`：停止按钮不再调用 `thread.terminate()`，改为 `request_stop()`
- `report_maker.py`：新增 `PipelineCancelled`、`check_cancelled()` 和 `cancel_check` 传递
- `report_maker.py`：`run_pipeline()`、`process_single_batch()`、`process_excel()`、`excel_to_pdf()`、外部链接处理支持取消检查
- `run_queue.py`：路径配置改为自动根据 `Path(__file__).resolve().parent` 推导
- `run_queue.py`：`QUEUE_DIR`、`DONE_DIR`、`REPORT_MAKER`、`LOG_DIR` 不再写死项目绝对路径
- `run_queue.py`：`PYTHON` 改为 `sys.executable`，使用当前运行 `run_queue.py` 的 Python 环境

### 使用方式
- 用法不变：进入项目目录后运行 `python run_queue.py ...`
- 项目目录移动后无需手动修改 `run_queue.py` 路径配置

### 验证
- `python -m py_compile report_maker.py gui_app.py run_queue.py utils.py` 通过
- 模拟立即停止：`run_pipeline()` 返回 `status=cancelled`
- 模拟 Excel 已打开后停止：确认调用 `Workbook.Close(SaveChanges=False)`、`Excel.Quit()`、`_kill_excel_process()`
- `python run_queue.py --dry-run --yes` 通过，可自动识别当前项目目录下的 `queue`

### 文档更新
- README.md 增加 Q6 常见问题说明
- PROJECT_LOG.txt 补充 2026-04-08 至 2026-04-21 前期开发记录
- PROJECT_LOG.txt 追加 2026-06-02 Hotfix 和安全停止优化记录
- README.md / PROJECT_LOG.txt 记录 `run_queue.py` 路径自动化改动

---

## v3.10 (2026-04-30) - Office/WPS下拉框

### 新增功能
- **Office/WPS下拉框**：gui_app.py参数区新增 `QComboBox`，运行时切换Excel类型
  - Office → `EXCEL_VISIBLE=False`（窗口隐藏，静默运行）
  - WPS → `EXCEL_VISIBLE=True`（窗口闪现一下）

### 问题修复

**问题2：WPS提示无法打开 `~$汇总表.xlsx`**
- 原因：`find_excel_files()`只按扩展名和“汇总”关键字筛选，未排除 WPS/Excel 生成的 `~$` 临时锁文件
- 现象：目录里残留 `~$汇总表.xlsx` 时，程序会误把它当作正式汇总表打开，WPS 弹出“无法打开文件”
- 解决：`find_excel_files()`增加 `f.startswith('~$')` 过滤，跳过临时文件

**问题1：选WPS后窗口仍隐藏**
- 原因：`run_pipeline()`的`excel_type=None`时没有兜底读取环境变量
- 解决：`run_pipeline()`加兜底逻辑，从`os.environ.get("RPT_EXCEL_TYPE", "office")`读取

### 代码修改

**gui_app.py：**
1. 删掉顶层`from report_maker import`（延迟导入，避免COM初始化冲突）
2. 新增`QComboBox`导入
3. 参数区加Office/WPS下拉框（第273-279行）
4. `ProcessThread`增加`excel_type`参数，启动时设`os.environ["RPT_EXCEL_TYPE"]`
5. `start_process`读取下拉框值传给`ProcessThread`
6. `on_finished`/`stop_process`重新启用下拉框
7. `_load_default_path`用`QTimer.singleShot(0, ...)`延迟执行

**report_maker.py：**
1. `run_pipeline()`增加`excel_type`参数
2. `excel_type=None`时从环境变量兜底读取
3. 两处`EXCEL_VISIBLE`设置点加DEBUG打印（Step5第233行、Step6第747行）
4. `find_excel_files()`过滤 `~$` 开头的 WPS/Excel 临时锁文件，避免误打开 `~$汇总表.xlsx`

### 打包命令
```bash
# 命令行方式
python -m PyInstaller --onefile --clean --hidden-import win32timezone --hidden-import win32crypt --hidden-import pywinauto --hidden-import pywinauto.keyboard --hidden-import pywinauto.timings --hidden-import pywinauto.application --hidden-import pywinauto.findwindows --name "Excel报表工具_GUI版_v3.10" gui_app.py
```

### exe相关文件
| 文件                         | 作用 |
|----------------------------|------|
| `gui_app.py`               | GUI打包入口，PyQt5图形界面 |
| `report_maker.py`          | 核心逻辑，被gui_app延迟导入 |
| `run_queue.py`             | CLI批量调度入口 |
| `utils.py`                 | 共享工具模块 |
| `dist\Excel报表工具_GUI版_v3.10` | 最终exe（64.4 MB） |

### 已知问题
- **DPI自适应**：125%/225%等不同缩放下界面显示大小不同，用户放弃实现

---

## v3.9 (2026-04-29) - 打包exe + PyQt5图形界面

### 新增功能
- **PyQt5 图形界面**（gui_app.py）：双击 exe 直接弹出界面，无需命令行，点按钮操作
- **exe 打包**：两种打包形态
  - `Excel报表工具_GUI版_v3.9.exe`（Office版，64.4 MB）
  - `Excel报表工具_命令行版_v3.9.exe`（命令行版，28 MB）
- **延迟导入机制**：解决 PyQt5 与 win32com 的 COM 初始化冲突问题

### 代码修改

**gui_app.py**（7处修改）：
1. 删掉顶层 `from report_maker import`（延迟导入）
2. `_load_default_path()` 加延迟导入
3. `start_process()` 加延迟导入
4. `ProcessThread.run()` 加延迟导入
5. 恢复 `QFileDialog`（替代 ctypes 原生对话框）
6. 恢复 `QMessageBox`（替代 tkinter messagebox）
7. 恢复 `QFileDialog` import

**report_maker.py**（5处修改）：
1. 新增 `find_latest_folder()` 函数（GUI 自动检测最新一期）
2. 新增 `run_pipeline()` 函数（GUI 专用入口）
3. `find_all_excel_files()` 加 `and not f.startswith('~$')`（过滤 Excel 临时文件）
4. `pywinauto.Desktop` 初始化加 `try/except` 保护
5. `_click_excel_update_button()` 加 `_desktop is None` 检查

**utils.py**（1处修改）：
1. `sys.stdout.isatty()` → `sys.stdout is not None and sys.stdout.isatty()`（打包后 sys.stdout 为 None）

### 打包命令

```bash
pyinstaller --onefile --collect-all pywinauto --hidden-import win32timezone --hidden-import win32crypt --name "Excel报表工具_GUI版_v3.8" gui_app.py
```

### exe 相关文件

| 文件 | 作用 |
|------|------|
| `gui_app.py` | GUI 打包入口，PyQt5 图形界面 |
| `report_maker.py` | 核心逻辑，被 gui_app 延迟导入 |
| `run_queue.py` | CLI 批量调度入口 |
| `utils.py` | 共享工具模块 |
| `dist\Excel报表工具_GUI版_v3.8.exe` | 最终产物（Office 版） |
| `dist\Excel报表工具_命令行版_v3.8.exe` | 命令行版 exe |

### 已知限制
- 当前 exe 为 **Office 版本**。如需 WPS 版本：将 `report_maker.py` 中 `EXCEL_TYPE = "office"` 改为 `EXCEL_TYPE = "wps"`，重新打包即可

---

## v3.8 (2026-04-28) - Step4全部用COM完成，解决I列超范围问题

### 问题描述
- **现象**：代码生成的汇总表I列（本次变量）超出理论范围±0.21，手工做的正常
- **理论范围**：I = -RANDBETWEEN(-21,12)/100，所以范围是 -0.21 到 +0.12
- **代码生成的值**：出现了0.29、-0.32等超出范围的值

### 问题分析
1. **I列公式**：I = (L-M)*1000 = (D-E)*1000，E = F数值 = RANDBETWEEN(-21,12)/100000 + D
2. **手工操作时**：复制E到D会触发F列重新计算（RANDBETWEEN随机一次）
3. **代码操作时**：openpyxl的赋值操作不会触发Excel重新计算F列，导致E列写入的是F的旧缓存值

### 解决方案
**全部用COM完成Step4**：
- 封面修改和日期列处理全部用COM完成，一次保存
- D=E操作在COM里执行，会触发F列重新计算
- 读取F列最新值写入E列

### 具体改动
1. **process_excel函数重写**：
   - 封面修改改用COM（之前用openpyxl会被COM保存覆盖）
   - 日期列处理全部用COM
   - 不再使用openpyxl的load_workbook

2. **新增days参数**：
   - `process_excel(filepath, target_date, issue_num, weather, days=1)`
   - 在处理三列数据之前，先更新第3列公式 =D1+days

3. **列处理逻辑（COM）**：
   ```
   1. COM打开文件 → Excel自动计算RANDBETWEEN
   2. 等待2秒 → 等Excel完成初始化计算
   3. ws.Calculate() → 强制计算一次
   4. 读取E列值到内存
   5. 把E列值写入D列 → 触发F列重新计算（关键！）
   6. 等待1秒
   7. 修改E列第1行（表头日期）
   8. 再次Calculate() → 确保F最新
   9. 读取F列（最新值）写入E列
   10. 保存关闭
   ```

### 待解决问题
- **TARGET_SHEETS名称不匹配**：实际Excel文件中的sheet名称与定义不匹配，导致部分sheet被跳过

---

## v3.7 (2026-04-22) - 代码质量优化 + 配置区重组

### 新增共享模块
- **新建 `utils.py`**：抽取 IDE 环境检测 + 颜色函数的共享代码
- `report_maker.py` 和 `run_queue.py` 均从 `utils.py` 导入，消除重复代码

### 代码重构
1. **`with_retry` 装饰器**：天气 API 三个函数的重试逻辑抽取为通用装饰器，减少约 45 行重复代码
   - `fetch_open_meteo_weather()`
   - `fetch_open_meteo_historical()`
   - `_fetch_seniverse_weather()`
2. **`main()` 函数拆分**：拆为 `parse_user_input()` + `process_single_batch()` + `main()` 三个函数，结构更清晰
3. **配置区重组**：
   - `report_maker.py`：顶部新增「使用者配置区」，所有可修改配置（路径、EXCEL_TYPE、API Key 等）集中在文件最前面，附中文注释
   - `run_queue.py`：同步重组配置区

### Bug 修复
- `process_excel` 函数中 `wb_data` 未定义 bug：在 `try` 之前先定义 `wb_data = None`

### 文件变化
| 文件 | 变化 |
|------|------|
| `report_maker.py` | ~950行，净减少约30行重复代码 |
| `run_queue.py` | ~430行，净减少约10行重复代码 |
| `utils.py` | 新增（约50行） |

---

## v3.6 (2026-04-22) - WPS 兼容性问题修复

### 问题描述
- **现象**：在 WPS 环境下，`excel.Visible = False` 时 pywinauto 无法点击"更新"对话框按钮
- **原因**：隐藏窗口的消息机制有问题——WPS 对话框依赖前台窗口，隐藏时 C 键消息被主窗口吃掉
- **Office 正常**：`excel.Visible = False` 时 Office 能正确处理，WPS 不行

### 解决方案
新增 `EXCEL_TYPE` + `EXCEL_VISIBLE` 配置，根据 Excel 类型自动设置 Visible：

```python
EXCEL_TYPE = "office"   # "wps" 或 "office"
EXCEL_VISIBLE = (EXCEL_TYPE.lower() == "wps")  # WPS必须True，Office用False
```

| EXCEL_TYPE | EXCEL_VISIBLE | 效果 |
|------------|--------------|------|
| `"office"` | `False` | 静默运行（无窗口） |
| `"wps"` | `True` | 窗口会闪一下，但 pywinauto 能正常点击 |

### 使用方式
顶部配置区直接修改一行：
```python
EXCEL_TYPE = "wps"   # WPS用户
EXCEL_TYPE = "office"  # Office用户
```

---

## v3.5 (2026-04-20) - Excel静默运行 + 项目整理

### Excel 无窗口运行
- **`excel.Visible = False`**：完全隐藏 Excel 窗口，不弹界面
- **`excel.ScreenUpdating = False`**：关闭屏幕更新，提升性能
- **`excel.DisplayAlerts = False`**：禁止弹窗警告
- pywinauto 仍会自动点击隐藏的对话框按钮，外部链接正常更新

### 项目结构整理
- 根目录重命名：`project` → `rpt-maker`
- `dispatcher.py` → `run_queue.py`
- `archive/` → `.archive/`
- `logs/` → `.logs/`
- `待处理队列/` → `queue/`
- `已完成/` → `done/`
- 项目迁移至：`F:\myproject\rpt-maker`

---

## v3.4 (2026-04-17) - 自动化断点 + 健壮性增强 + 用户体验优化

### 新增功能

**run_queue.py：**
- **`--notify / -n`**：处理完成后发 Windows Toast 通知（优先用 win10toast 库，fallback 到 PowerShell）
- **`--retry-failed`**：读取 `.logs/failed_projects.txt`，只重试历史失败的工程（需配合 `--yes`)
- **`--dry-run`**：只显示要做什么，不实际执行（用于预览）
- **失败记录持久化**：成功后从 `failed_projects.txt` 移除，失败时加入，重试命令：`python run_queue.py --retry-failed --yes`

**report_maker.py 健壮性增强：**
- **Step5 文件级重试**：每个 Excel 文件最多重试 3 次，指数退避（1s/2s/4s）
- **PDF 转换重试**：PDF 生成失败自动重试最多 3 次，指数退避
- **幂等性保护**：`copy_directory` 加 `force` 参数，pipeline 模式下自动覆盖不询问
- **天气 API 指数退避**：心知天气、Open-Meteo Forecast、Historical 三个 API 各最多重试 3 次（1s/2s/4s）
- **重试配置常量**：`MAX_RETRIES=3`，`RETRY_BACKOFF=[1, 2, 4]`

**report_maker.py 用户体验：**
- **Windows Console 进度条**：用 Windows API 直接写屏幕，绕过 stdout 管道缓冲，TTY 下实时更新
- **成功/失败颜色**：绿色`[成功]`、红色`[失败]`、青色`[完成]`、黄色`[警告]`
- **可选 tqdm 进度条**：已安装 tqdm 时可用（未安装自动退化，不报错）

**run_queue.py 输出优化：**
- **实时逐行输出**：用 `iter(process.stdout.readline, '')` 替代 `process.communicate()`，子进程输出不再延迟
- **编码修复**：子进程环境变量加 PYTHONIOENCODING=utf-8 / PYTHONUTF8=1，encoding=utf-8，彻底解决跨编码环境乱码问题

### 问题修复（2026-04-17 下午）

**PyCharm 颜色兼容：**
- **问题**：PyCharm 运行 dispatcher 时显示 `[0m` 等 ANSI 转义序列
- **原因**：`PYCHARM_HOSTED` 环境变量在 PyCharm run configuration 中未设置
- **修复**：通过 `sys.executable` 路径检测 PyCharm 环境（最可靠），同时检测 `PYCHARM_HOSTED`、`PYCHARM`、`VSCODE_PID`、`VSCODE_INJECTION`
- **IDE 环境自动禁用颜色**：检测到 IDE 环境时强制关闭颜色输出

**天气 API 逻辑优化：**
- **问题**：`days_diff > 7` 时仍调用心知天气，导致逻辑混乱
- **修复**：`fetch_api_weather()` 中 `days_diff > 7` 时直接调用 Open-Meteo
- **日志优化**：移除 `get_weather()` 中误导性的"心知天气返回"日志

### 代码优化（2026-04-17 傍晚）
- 移除冗余 import：`locale`（两个文件均未使用）
- 移除冗余 `sys` 重复导入
- `TARGET_SHEETS` 从 list 改为 set，提升查找效率 O(n) → O(1)
- 移除未使用的 `ctypes.wintypes` 导入

### 使用示例

```bash
# 完全无人值守 + 通知
python run_queue.py --yes --notify --continue-on-error
python run_queue.py -y -n -c  # 简写

# 重试失败工程
python run_queue.py --retry-failed --yes

# 预览（不执行）
python run_queue.py --dry-run
```

---

## v3.3 (2026-04-16) - 封面修改逻辑重构

### 新增功能
- 封面搜索范围优化：`min_col=5, max_col=40`，更精准定位封面区域
- 支持数值类型期数单元格：`str(cell.value)` 统一转字符串处理，兼容数值类型（如 AC6 = 2）
- 天气前缀智能处理：
  - 有"天气："前缀：替换前缀后内容（如 天气：阴 → 天气：晴）
  - 无前缀：直接替换关键词本身（如 阴 → 晴）
- 去掉温度相关所有代码，封面只改期数和天气
- 增加调试打印，便于排查封面修改问题

### 代码优化
- 提取共享常量：`WMO_CODE_MAP` 和 `JINAN_LAT/LON` 提到模块顶部，消除重复
- 清理顶部 import：`urllib.request` 和 `json` 从函数内移至文件顶部
- 简化资源管理：`process_excel` 去掉冗余的 `wb_data = None` 初始化
- 稳定增强：`excel_to_pdf` 去掉多余的杀进程调用（Dispatch 新实例不需要）

### 项目结构
- 创建 `.archive/` 目录，存放历史版本
- 创建 `.logs/` 目录，存放调度器日志
- 重命名 `run_all_projects.py` → `run_queue.py`
- 清理冗余文件，项目结构更清晰

---

## v3.2 (2026-04-16) - 多工程调度器 + 三重天气系统

### 新增功能
- **多工程调度器** (`run_all_projects.py`)：
  - 支持队列处理：`F:\myproject\rpt-maker\queue\`
  - 每个工程文件夹下放 `config.txt`（格式如 `5,1,2,3,1,4`）
  - 处理完自动移到 `done/` 子目录
  - 每个工程有独立日志文件（`.logs/工程名_时间.log`）
- **三重天气API系统**：
  - 主用：心知天气（Key: `S-db8PJBJo-bZ0rb0`）
  - 备用：Open-Meteo Forecast API（未来天气）
  - 历史：Open-Meteo Historical API（过去日期，无需 API Key）
  - 获取策略：
    - `days_diff < 0`：Historical API 查真实历史
    - `0 <= days_diff <= 16`：心知天气，失败用 Forecast
    - `days_diff > 16`：季节模拟
- **环境变量支持**：`REPORT_MAKER_PARENT` 动态指定工程目录

### 修复问题
- 修复天气 API 错误：`UnboundLocalError: cannot access local variable 'temperature'`
- 修正函数签名和返回值

---

## v3.1 (2026-04-15) - Step5 全自动化 + 13项代码优化

### 核心突破
- **Step5 全自动化**：pywinauto + COM 混合方案成功
  - 7个纯 COM 方案失败后找到根本原因：Windows COM 单例机制
  - 最终方案：COM 负责打开/保存，pywinauto 点击"更新(&U)"按钮
  - 关键参数：`UpdateLinks=2` 强制弹出对话框，`DisplayAlerts=False` 对它无效

### 13项代码优化
1. 删掉 `_close_activation_wizard()`（死代码）
2. 简化 `open_and_update_links` 的 except 块
3. 清理 `renamed` 变量（冗余）
4. 简化 main 循环中间变量
5. 去掉 winreg 和 pywinauto 的 try/except 包裹
6. 删掉 main() 里的 if winreg is None 检查（死代码）
7. `open_and_update_links` 重试逻辑重构（for 循环→try/except）
8. `_click_excel_update_button` 的 Desktop 对象缓存到模块级 `_desktop`
9. `fetch_api_weather` 去掉 abs()
10. `get_year_from_excel` 加 try/finally
11. `process_excel` 里 wb_data 加 try/finally
12. 输入格式从多次 input 改成一次性逗号分隔
13. 调试打印加到 _do() 里（已注释备用）

### 交互改进
- 输入格式：从多次 `input()` 改为一次性逗号分隔（如 `5,1,2,3,1,4`）
- 支持外部自动化：`echo 5,1,2,3,1,4 | python report_maker.py`

---

## v3.0 (2026-04-14) - 批量处理框架完成

### 核心功能
- 六步主流程全部实现：
  1. 复制目录（期数+1，日期+自定义天数）
  2. 删除目录下所有 PDF 文件
  3. 修改 Excel 封面（期数、天气，温度）
  4. 处理 Excel 日期列（第2列→第1列，第3列→第2列，只写数值）
  5. 自动处理外部引用（注册表+自动点击更新）
  6. Excel 转 PDF（COM 接口）

### 技术突破
- 发现 openpyxl 写公式不写值的问题，用 `data_only=True` 补救
- 发现 COM 单例问题，开始研究 pywinauto 方案
- 7个 Step5 自动化版本全部失败，找到根本原因

---

## v2.0 (2026-04-13) - 单期版完善 + 天气API集成

### 主要改进
- 天气 API 集成：心知天气（济南）
- 天气关键词匹配：从精确匹配改为关键词列表遍历
- 温度处理：用正则 `r'\d+°C'` 匹配温度单元格
- 日期格兼容：`is_date_cell()` 同时检查 datetime 对象和字符串格式

### 发现问题
- openpyxl 写公式不写值的问题首次发现
- xlwings 和 openpyxl 混用导致 COM 冲突
- 开始研究 Step5 自动化方案

---

## v1.0 (2026-04-08) - 初始版本

### 基础功能
- 单期处理框架
- 基础日期处理逻辑
- Excel 转 PDF 功能
- 期数和日期递增逻辑

### 技术栈
- openpyxl 读写 Excel
- win32com 操作 Excel COM
- 心知天气 API 初步集成

---

## 重要教训总结

1. **openpyxl 写公式不写值**：必须用 `data_only=True` 补救
2. **混用库的 COM 冲突**：openpyxl 和 xlwings 不要同时用
3. **COM 单例问题**：Windows 的 `Dispatch()` 无法创建全新实例，必须配合 GUI 自动化
4. **UpdateLinks=2 对 DisplayAlerts=False 无效**：对话框一定会弹，必须 pywinauto 点击
5. **调试打印的价值**：`print(f"value={repr(cell.value)}")` 直接定位问题
6. **kill 进程时机**：必须在 Dispatch 之前，且要 sleep(1.5) 保证完全退出
7. **重试不是万能药**：每个文件独立杀进程比重试4次更有效
8. **资源管理**：openpyxl Workbook 不支持 with 语句，必须用 try/finally

---

## 文件对应关系（历史沿革）

| 原始文件名 | 新文件名 | 版本 | 状态 |
|-----------|----------|------|------|
| test.py | report_maker.single.py | v1.0 | 存档 |
| run_all.py | report_maker.manual.py | v2.0 | 存档 |
| run_all1.py | report_maker.py | v3.0+ | 当前 |
| run_all1_backup.py | report_maker.backup.py | v3.1 | 存档 |
| run_all_projects.py | run_queue.py | v3.2 | 当前 |
| excel_to_pdf.py | - | - | 已集成 |
| _runner.py | - | - | 已删除 |

---

---

## 📊 项目状态

**当前版本**：**v3.11.3**（WPS PDF export guard + #### column-width fix）
**最后更新**：2026-07-22
**维护者**：朱无理
**Python版本**：3.8+
**操作系统**：Windows 10/11
**依赖库**：openpyxl, pywin32, pywinauto, requests, PyQt5

## 📁 当前文件结构

```
F:\myproject\rpt-maker\
├── report_maker.py          # 主程序 v3.11（含run_pipeline入口）
├── run_queue.py             # CLI批量调度器
├── gui_app.py               # GUI打包入口 v3.11（PyQt5）
├── utils.py                 # 共享工具模块
├── dist\                    # exe打包输出目录
│   └── ExcelReports_GUI_v3.9.exe   # 最终exe（64.4 MB）
├── queue\                   # 待处理工程队列
│   └── done\               # 已完成工程
├── CHANGELOG.md            # 本文件
├── README.md               # 使用说明
└── PROJECT_LOG.txt         # 详细开发记录
```

*文档最后更新：2026-07-22*
