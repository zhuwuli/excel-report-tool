# Excel汇总表自动处理工具 v3.11.3

> 自动把工程监测数据做成 PDF 报表，六步流程一键完成。

---

## 快速开始

### 双击 exe 运行（最简单）

1. 双击 `dist\ExcelReportTool_v3.11.3_Windows_x64.exe`
2. 选择 Excel 类型（Microsoft Office 或 WPS Office）
3. 点击"选择文件夹"，找到工程文件夹
4. 填总期数和各期间隔天数
5. 点击"开始处理"

> Office 模式窗口隐藏静默运行，WPS 模式窗口会闪现一下，均能正常自动化处理。
> v3.11 起推荐使用 `--windowed` 无黑框 exe，运行错误会直接显示在 GUI 实时日志中。

### 工程文件夹准备

1. 把工程文件夹放入 `queue\` 文件夹
2. 工程文件夹里必须有 `config.txt` 文件

**config.txt 怎么写？**
```
总期数,第1期间隔,第2期间隔,第3期间隔,...
```

| config.txt内容 | 含义 |
|---------------|------|
| `3,1,1,1` | 3期，每期间隔1天 |
| `5,2,2,2,2,2` | 5期，每期间隔2天 |
| `1,0` | 1期（当天），间隔0天 |

### 取结果

处理完成后工程文件夹自动移到 `queue\done\`，直接去那里拿 PDF 结果。

---

## 日常使用详解

### 三种运行方式

**方式一：直接双击 exe（最简单，推荐）**

1. 进入 `dist\` 文件夹
2. 双击 `ExcelReportTool_v3.11.3_Windows_x64.exe`
3. 弹出图形界面
4. 选择 Excel 类型
5. 点击"选择文件夹" → 找到工程文件夹
6. 填总期数（如 `3`）和各期间隔天数（如 `1,1,1`）（英文逗号）
7. 点击"开始处理"

**方式二：命令行**

1. 打开命令提示符（Win + R → 输入 `cmd` → 回车）
2. 进入项目目录：`cd 你存放项目的路径`
3. 运行：`python run_queue.py -y -n --continue-on-error`

命令解释：
- `-y` = 自动确认，不用问你了
- `-n` = 完成后发通知
- `-c` = 某个失败了也继续处理下一个

**方式三：PyCharm（如果你装了 PyCharm）**

1. 用 PyCharm 打开项目文件夹
2. 右键点击 `run_queue.py`
3. 选择 "Run 'run_queue'"

### 等待完成

你会看到类似这样的输出：
```
[9:20:15] 开始处理：工程A
[9:20:16] Step1: 复制目录...
[9:20:17] Step2: 删除PDF...
[9:20:18] Step3: 修改封面（期数、天气）...
[9:20:19] Step4: 处理日期列...
[9:20:20] Step5: 处理外部引用...
[9:20:21] Step6: 转换为PDF...
[9:20:25] ✓ 工程A 处理完成！
```

**等待时间：** 每个工程大约 1-5 分钟，取决于数据量。

### 取结果

处理完成后：
1. 工程文件夹会自动从 `queue` 移到 `queue\done\`
2. PDF 文件在工程文件夹里

---

## 项目文件夹结构

```
rpt-maker\
├── report_maker.py       # 主程序（六步核心逻辑）
├── run_queue.py          # CLI批量调度器
├── gui_app.py            # GUI打包入口（PyQt5）
├── utils.py              # 共享工具模块
├── dist\                 # exe输出目录
│   └── ExcelReportTool_v3.11.3_Windows_x64.exe  # 双击exe直接用（无黑框GUI）
├── queue\                # 把要处理的工程文件夹放这里！
│   └── done\            # 处理完的工程会自动移到这里
├── .logs\               # 日志文件夹（记录运行过程）
├── .archive\            # 历史版本存档（旧版 report_maker / gui_app 备份）
├── .idea\               # PyCharm配置
├── .venv\               # Python虚拟环境
├── README.md            # 说明文档
├── CHANGELOG.md         # 版本记录
└── PROJECT_LOG.txt      # 开发日志
```

**你只需要关心 3 个地方：**
1. **`queue\` 文件夹** - 放要处理的工程
2. **`queue\done\` 文件夹** - 处理完的工程在这儿
3. **`dist\ExcelReportTool_v3.11.3_Windows_x64.exe`** - 双击运行（无黑框GUI）

---

## 常见问题

### Q1: 报错"找不到模块"
**原因：** 没安装库
**解决：**
```bash
pip install openpyxl pywin32 pywinauto requests PyQt5 pytz
```

### Q2: 报错"系统找不到指定的路径"
**原因：** 工程目录、`queue` 文件夹或 `report_maker.py` 缺失
**解决：**
1. 确保在项目目录里运行 `python run_queue.py ...`
2. 确保项目目录下有 `queue\` 文件夹
3. 确保 `report_maker.py` 和 `run_queue.py` 在同一个项目目录里

> v3.10 Hotfix 后，`run_queue.py` 会自动根据脚本所在目录推导路径，一般不需要手动改路径。

### Q3: 报错"Python不是内部或外部命令"
**原因：** Python 没装好，或者环境变量没配
**解决：**
1. 去 https://python.org 下载 Python
2. 安装时**务必勾选** "Add Python to PATH"（很重要！）
3. 安装完成后重新打开命令提示符

### Q4: 运行后卡住了
**原因：** Step5 可能正在等待 Excel/WPS 启动、外部链接更新或公式计算。
**解决：** 当前版本已改为智能等待：Office 更新按钮最多等待 6 秒，WPS 最多等待 8 秒，公式计算最多等待 5 秒；状态就绪后会立即继续。若单个文件长时间无新日志，可点击 GUI 的“停止”进行安全退出后重试。

### Q5: 天气显示"模拟天气"
**原因：** 天气 API 查不到（日期太远或网络问题）
**解决：** 这是正常的！会用估算的天气，不影响结果。

### Q6: WPS提示"无法打开文件"，文件名是 `~$汇总表.xlsx`
**原因：** `~$` 开头的是 WPS/Excel 临时锁文件，不是真正的 Excel 表格。旧版本会把它误识别成汇总表并尝试打开。
**解决：**
1. 关闭所有 WPS / Excel 窗口
2. 删除工程目录里所有 `~$` 开头的文件
3. 使用已修复版本重新运行

**代码修复：** `report_maker.py` 的 `find_excel_files()` 已增加过滤逻辑，会自动跳过 `~$` 临时文件；复制新一期目录时也会自动过滤 `~$` 临时锁文件，避免继续带入新目录。

### Q7: 一开始就提示"运行前自检失败"
**原因：** 程序在正式处理前发现基础条件不满足，例如工程路径不存在、找不到最新一期文件夹、最新一期里没有汇总表，或 Excel 自动化依赖不可用。
**解决：**
1. 检查选择的目录是否是工程目录或某一期目录
2. 确认目录里有形如 `第43期 05.30` 的文件夹
3. 确认最新一期里有 `汇总表.xlsx`
4. 如果提示依赖不可用，重新安装：`pip install pywin32 pywinauto`

---

## 快速命令参考

| 要做什么 | 命令 |
|---------|------|
| 处理队列里的所有工程 | `python run_queue.py -y -n --continue-on-error` |
| 预览要处理哪些（不实际执行） | `python run_queue.py --dry-run` |
| 重试之前失败的工程 | `python run_queue.py --retry-failed --yes` |
| 查看帮助 | `python run_queue.py --help` |

---

================================================================================
以下为技术文档（进阶内容，新手无需阅读）
================================================================================

<details>
<summary><strong>📗 技术文档 — 点击展开</strong></summary>

---

# 📗 技术文档（给有基础的人看）

---

## 🔧 六步处理流程

### Step 1：复制目录
- 期数+1（如：第87期 → 第88期）
- 日期+自定义天数（如：04.09 → 04.10）
- 保留所有文件结构

### Step 2：删除PDF文件
- 清理新目录中的所有 PDF 文件（避免重复）

### Step 3：修改Excel封面
- **期数**：查找"第X期"单元格，更新为新期数
- **天气**：根据目标日期自动获取济南天气
  - 支持"天气：晴"格式（保留前缀）
  - 支持"晴"格式（无前缀）
  - 搜索范围：1-20行，5-40列
- **兼容处理**：支持数值类型期数单元格（如 AC6 = 2）

### Step 4：处理Excel日期列
- **第2列 → 第1列**：数据迁移
- **第3列 → 第2列**：只写数值（不写公式）
- **第2列表头**：更新为目标日期
- **兼容处理**：同时支持 datetime 对象和字符串格式

### Step 5：自动处理外部引用（核心难点）

**目标**：替代人工点击"是否更新链接"对话框

**技术方案**：pywinauto + COM 混合

**完整流程**：
```
_kill_excel_process()              → 杀残进程
↓
Dispatch 新 Excel 实例（Visible=EXCEL_VISIBLE） → 拿到干净实例
↓
打开文件（UpdateLinks=2）           → 触发链接更新对话框
↓
等 2 秒                             → 等 Excel 完全启动
↓
等 2 秒                             → 等对话框弹出来
↓
pywinauto 点"更新(&U)"            → 自动化点击
↓
sleep(1.0)                         → 等 Excel 完成链接更新和计算
↓
wb.Save() + wb.Close()             → 保存关闭
↓
excel.Quit()                        → 关闭实例
↓
_kill_excel_process()              → 再杀一遍
```

**WPS 兼容性**：
- **Office**：`EXCEL_VISIBLE = False`（窗口隐藏，静默运行）
- **WPS**：`EXCEL_VISIBLE = True`（窗口会闪一下，但 pywinauto 能正常点击）
- 只需修改顶部的 `EXCEL_TYPE = "wps"` 或 `"office"`

**关键参数**：
- `excel.Visible = EXCEL_VISIBLE`：根据类型自动设置
- `excel.ScreenUpdating = False`：关闭屏幕更新，提升性能
- `excel.DisplayAlerts = False`：禁止弹窗（但链接更新对话框仍会弹）
- `UpdateLinks=2`：强制弹出"是否更新链接"对话框
- `wait_until_passes(12, 0.5)`：最多等12秒，每0.5秒重试一次
- 每个文件独立杀进程，保证 COM 实例 freshness

### Step 6：Excel转PDF
- 自动将"第XX期.xlsx"转换为 PDF
- 保持所有格式和布局
- **失败自动重试**：最多3次，指数退避（1s/2s/4s）
- **静默模式**：PDF转换时 Excel 窗口同样隐藏

---

## 📊 目标工作表

工具只处理以下 **6个** 工作表，其他自动跳过：
1. 垂直位移
2. 坑外水位
3. 测斜
4. 水平位移
5. 竖向位移
6. 轴力

---

## 🔄 健壮性与重试

| 操作 | 重试次数 | 退避策略 |
|------|---------|---------|
| Step5 外部引用处理 | 3次/文件 | 1s → 2s → 4s |
| PDF 转换 | 3次/文件 | 1s → 2s → 4s |
| 心知天气 API | 3次 | 1s → 2s → 4s |
| Open-Meteo Forecast API | 3次 | 1s → 2s → 4s |
| Open-Meteo Historical API | 3次 | 1s → 2s → 4s |

**幂等性保护**：`copy_directory` 默认 `force=True`，覆盖已有目录不询问。

**重试配置常量**：
- `MAX_RETRIES = 3`
- `RETRY_BACKOFF = [1, 2, 4]`

---

## 💡 用户体验增强

**无窗口运行（Office）**：Excel 全程在后台隐藏运行，不弹 Excel 界面（`Visible=False` + `ScreenUpdating=False`）。

**WPS 兼容**：WPS 用户设置 `EXCEL_TYPE = "wps"`，窗口会闪一下但自动化正常工作。

**进度条**：Step5 处理时显示进度条，TTY 下实时更新（Windows Console API 实现）。

**颜色输出**：成功/失败/完成/警告 分别显示为绿/红/青/黄色。

**实时输出**：run_queue.py 用逐行读取替代管道缓冲，子进程输出不再延迟到结束。

**IDE兼容**：检测到 PyCharm/VSCode 环境时自动禁用颜色输出。

**可选 tqdm**：已安装 tqdm 时可用，未安装自动退化，不报错。

---

## 🌤️ 天气系统

### 三重获取方案

| 日期范围 | 使用API | 说明 |
|---------|---------|------|
| 过去日期 | Open-Meteo Historical | 查询历史真实天气，无需API Key |
| 今天 ~ 16天后 | 心知天气（主用）/ Open-Meteo（备用） | 天气预报 |
| 16天以后 | 模拟天气 | 根据季节估算 |

### API配置
- **心知天气 Key**: `S-db8PJBJo-bZ0rb0`
- **济南坐标**: 36.6512°N, 117.1205°E

### 获取策略
- `days_diff < 0`：Historical API 查真实历史
- `0 <= days_diff <= 16`：心知天气，失败用 Forecast
- `days_diff > 16`：季节模拟

### WMO天气码映射
心知天气和 Open-Meteo 共用 WMO_CODE_MAP，将国际气象码转换为中文天气关键词。

---

## ⚙️ 环境配置

### Excel 类型配置
```python
# 顶部配置，根据你的 Excel 类型选择：
EXCEL_TYPE = "office"   # Microsoft Office（默认，静默运行）
EXCEL_TYPE = "wps"      # WPS（窗口会闪一下，但自动化正常）
```

### Python依赖
```bash
pip install openpyxl pywin32 pywinauto requests PyQt5 pytz
```

### 开发环境与 exe 打包

项目当前 `.venv` 使用 Python 3.8.6。打包前应确认终端只激活项目虚拟环境，并检查解释器路径：

```powershell
.\.venv\Scripts\Activate.ps1
python -c "import sys; print(sys.executable)"
```

输出应指向 `rpt-maker\.venv\Scripts\python.exe`。推荐使用：

```powershell
python -m PyInstaller --onefile --clean `
  --hidden-import win32timezone `
  --hidden-import win32crypt `
  --hidden-import pywinauto `
  --hidden-import pywinauto.keyboard `
  --hidden-import pywinauto.timings `
  --hidden-import pywinauto.application `
  --hidden-import pywinauto.findwindows `
  --name "ExcelReportTool_v3.11.3_Windows_x64" gui_app.py
```

不要在当前环境使用 `--collect-all pywinauto`。如果打包日志显示 `Python environment: F:\anaconda`，说明实际调用了 Anaconda 的 PyInstaller，可能把 pandas、scipy、sklearn、Jupyter 等无关依赖一起收集，导致 exe 从约 66 MB 增长到数百 MB。

### 系统要求
- Windows 10/11
- Microsoft Excel 2016+ 或 WPS
- Python 3.8+

### 环境变量
```bash
# 动态指定工程目录（用于调度器）
set REPORT_MAKER_PARENT=F:\myproject\rpt-maker\queue\工程A
```

---

## 🔍 调试技术

### 调试打印（已注释备用）
在 `open_and_update_links()` 的 `_do()` 函数内：
```python
# print(f"  [调试] 开始保存...")
wb.Save()
# print(f"  [调试] 保存完成，开始关闭...")
wb.Close()
# print(f"  [调试] 关闭完成，开始退出Excel...")
excel.Quit()
# print(f"  [调试] Excel已退出")
```

### 日志系统
多工程调度器自动生成日志，文件名格式：`工程名_YYYYMMDD_HHhMMmSSs.log`
```
F:\myproject\rpt-maker\.logs\
├── 工程A_20260416_16h44m27s.log
└── 工程A_20260416_16h50m43s.log
```

---

## 🐛 常见问题详解

### Q1: 封面修改失败
**症状**：期数或天气没有更新
**解决**：
1. 检查单元格是否为数值类型（如 `2` 而不是 `"第2期"`）
2. 检查搜索范围（1-20行，5-40列）
3. 启用调试打印查看查找过程

### Q2: Excel卡在"是否更新链接"对话框
**症状**：程序卡住，需要手动点击
**原因**：COM 单例问题，Dispatch() 返回已损坏的实例
**解决**：
1. 确保 `pywinauto` 已正确安装
2. 检查Excel窗口标题是否为"Microsoft Excel"
3. 按钮文本应为"更新(&U)"
4. pywinauto + COM 混合方案自动处理（窗口隐藏但按钮仍在）

### Q3: 天气获取失败
**症状**：显示"模拟天气"而不是真实天气
**解决**：
1. 检查网络连接
2. 确认心知天气API Key有效
3. 检查日期是否超出16天范围

### Q4: PDF转换失败
**症状**：Excel文件存在但PDF未生成
**解决**：
1. 确保Excel文件已保存
2. 检查COM权限
3. 确认Excel进程已完全关闭
4. 启用重试机制（最多3次）

</details>

<details>
<summary><strong>⚠️ 重要教训 — 点击展开</strong></summary>

---

## ⚠️ 重要教训（踩坑总结）

### 技术问题

1. **openpyxl 写公式不写值**：必须用 `data_only=True` 补救

2. **混用库的 COM 冲突**：openpyxl 和 xlwings 不要同时用

3. **COM 单例问题**：Windows 的 `Dispatch()` 无法创建全新实例，必须配合 GUI 自动化

4. **UpdateLinks=2 对 DisplayAlerts=False 无效**：对话框一定会弹，必须 pywinauto 点击

5. **调试打印的价值**：`print(f"value={repr(cell.value)}")` 直接定位问题

6. **kill 进程时机**：必须在 Dispatch 之前，且要 sleep(1.5) 保证完全退出

7. **重试不是万能药**：每个文件独立杀进程比重试4次更有效

8. **资源管理**：openpyxl Workbook 不支持 with 语句，必须用 try/finally

9. **winreg 是 Windows 内置**：不需要 try/except，缺库直接 import 失败比静默 None 更好

### 业务逻辑

10. **天气关键词匹配**：用关键词列表遍历匹配，而不是精确匹配单一字符串

11. **天气 API 的 days 参数从 0 开始**：心知天气 API 的第0天是今天

12. **日期写入格式**：写入字符串 `'2026/04/10'` 而不是 datetime 对象，兼容更好

### 代码设计

13. **函数设计**：单一职责优于大而全的函数，每个函数做好一件事

14. **代码审查**：发现问题后先确认是 bug 还是设计意图，再决定是否修改

15. **Windows subprocess**：所有输入在脚本启动时就灌进去，不支持真正的实时交互输入

### 代码评审优化建议（来自同事反馈）

16. **wb_data 未定义 bug**：process_excel 中 load_workbook 抛异常时 wb_data 未定义，finally 里会报 NameError。修复：在 try 之前先定义 wb_data = None

17. **天气函数重试逻辑重复**：fetch_open_meteo_weather、fetch_open_meteo_historical、_fetch_seniverse_weather 三个函数结构几乎相同，可抽成装饰器

18. **颜色检测可简化**：当前用 try/except，可简化为 `HAS_COLOR = sys.stdout.isatty() and not IDE_ENV`

19. **main() 函数过长**：可抽取 process_single_batch() 等子函数，提升可读性

</details>

<details>
<summary><strong>📝 版本历史 — 点击展开</strong></summary>

---

## 📝 版本历史
### v3.11.3 (2026-07-22) - WPS PDF export guard and #### column-width fix
- Added WPS-only PDF PageSetup guard before `ExportAsFixedFormat()`.
- WPS visible sheets now use A4, `Zoom = False`, and `FitToPagesWide = 1` before PDF export.
- Fixed one-page sheets in WPS mode use `FitToPagesTall = 1`.
- Data sheets keep natural vertical pagination, so manually added sheets or longer data can still produce more pages.
- Added pre-PDF `####` detection using COM `cell.Text`, not cell value search.
- `####` fix only changes column width. It does not change font size and does not use `ShrinkToFit`.
- Each affected column can widen up to 3 times. Each step uses the larger of `+2` or `*1.15`, capped at column width 30.
- Column-width changes are now saved back into the generated Excel workbook with `wb.Save()`.
- Logs now show found/resolved/unresolved counts, sheet/column, width changes, attempt count, and save status.
- Morning test builds: `ExcelReportTool_v3.11.1_Windows_x64`, `ExcelReportTool_v3.11.2_Windows_x64`.
- Current recommended build: `ExcelReportTool_v3.11.3_Windows_x64`.

### v3.11 (2026-07-03) - GUI 重设计 + 无黑框发布体验
- GUI 从原单列表单改为“工作台”布局：顶部品牌区、左侧任务配置、右侧运行状态与实时日志
- 界面版本号更新为 `v3.11`，推荐 exe 名称更新为 `ExcelReportTool_v3.11.3_Windows_x64.exe`
- `--windowed` 打包适配：`sys.stdout/sys.stderr` 为 `None` 时自动兜底，避免无控制台环境闪退
- 后台 `print()`、`stderr` 和异常 traceback 会转入 GUI 日志区，错误不再只出现在黑框
- `_kill_excel_process()` 调用 `taskkill` 时隐藏子进程窗口，避免运行中闪出 `taskkill` 黑框
- 实时日志增加颜色区分，成功、警告、失败和处理中信息更容易辨认
- 左侧配置区重排并修复小窗口下文字重叠、按钮压住内容的问题
- 保留原 Excel/WPS 自动化处理逻辑，本次主要是 GUI 与发布体验升级

### v3.10 Hotfix (2026-06-27) - STEP4/STEP5 性能优化

**STEP4 数据迁移优化**：
- Excel 单元格逐行 COM 读写改为 `Range.Value` 批量读写
- 保留“第3列为空时不覆盖第2列原值”的业务规则
- 备份：`.backup\20260626_164422\report_maker.before-step4-range.py`

**STEP5 智能等待优化**：
- 启动等待由固定 2 秒改为检测 `Excel.Ready`，至少 0.5 秒、最多 2 秒
- 无外部链接的工作簿直接跳过更新按钮等待
- Office/WPS 更新按钮分别最多轮询 6 秒/8 秒，找到后立即继续
- 更新后的固定 1 秒改为检测 `CalculationState`，至少 0.5 秒、最多 5 秒
- 文件间隔由 1 秒缩短为 0.2 秒，并新增单文件耗时日志
- 保留逐文件启动、保存、关闭和进程清理逻辑

**验证**：
- `python -m py_compile report_maker.py gui_app.py run_queue.py` 通过
- STEP4 Range 数据形状与空值保留辅助测试通过
- STEP5 就绪等待、计算等待、链接检测、按钮轮询、安全取消和完整模拟流程通过

---

### v3.10 Hotfix (2026-06-26) - 开发环境与打包修复

- 发现旧 `.venv` 指向已不存在的 `C:\Python314`，使用本机 Python 3.8.6 重建
- 新 `.venv` 已验证 PyQt5、openpyxl、win32com、pywinauto、requests、PyInstaller、tqdm 可导入
- 确认同时显示 `(.venv) (base)` 时，直接执行 `pyinstaller` 仍可能调用 `F:\anaconda\Scripts\pyinstaller.exe`
- 打包方式改为项目环境中的 `python -m PyInstaller`
- 移除推荐命令中的 `--collect-all pywinauto`，避免 Anaconda 无关依赖导致 exe 膨胀到约 600 MB

---

### v3.10 Hotfix (2026-06-26) - 鲁棒性增强

**新增增强**：
- WPS 模式下清理 `EXCEL.EXE` 的同时额外清理 `et.exe`
- 复制新一期目录时自动过滤 `~$` 开头的 Excel/WPS 临时锁文件
- 新增运行前轻量自检，提前检查路径、最新一期、汇总表、Excel 类型和自动化依赖

**验证**：
- `python -m py_compile report_maker.py gui_app.py run_queue.py utils.py` 通过
- 临时目录测试确认 `~$` 文件不会被复制到新一期
- 自检测试确认工程父目录和某一期目录均可识别最新一期

---

### v3.10 Hotfix (2026-06-02) - 临时文件过滤 + 安全停止 + 路径自动化

**问题修复**：
- 修复 WPS 提示无法打开 `~$汇总表.xlsx` 的问题
- `find_excel_files()` 过滤 `~$` 开头的 WPS/Excel 临时锁文件
- 修复 GUI 中途点击“停止”容易残留 Excel 进程的问题

**安全停止优化**：
- 停止按钮不再硬终止后台线程
- 改为请求安全停止，当前步骤收尾后再退出
- Excel COM 操作取消时会尽量关闭工作簿、退出 Excel 并清理进程

**路径自动化**：
- `run_queue.py` 自动根据当前脚本所在目录推导 `queue`、`.logs`、`report_maker.py`
- `PYTHON` 改为使用当前运行 `run_queue.py` 的 Python 环境
- 项目移动到其他目录后，不需要再手动修改 `run_queue.py` 顶部路径

**使用方式**：
- 用法不变，进入项目目录运行 `python run_queue.py ...`
- 预览：`python run_queue.py --dry-run --yes`
- 正式处理：`python run_queue.py -y -n -c`

**验证**：
- 语法检查通过
- 模拟立即停止返回 `cancelled`
- 模拟 Excel 已打开后停止，确认执行 `Close` / `Quit` / 清理进程
- `python run_queue.py --dry-run --yes` 通过

**文档更新**：
- 常见问题新增 Q6
- PROJECT_LOG.txt 补充前期开发记录和本次 Hotfix / 安全停止 / 路径自动化记录

---

### v3.10 (2026-04-30) - Office/WPS下拉框

**新增功能**：
- **Office/WPS下拉框**：gui_app.py 参数区新增下拉框，运行时切换 Excel 类型
**问题修复**：
- 选 WPS 后窗口仍隐藏的问题

---

### v3.9 (2026-04-29) - 打包exe + PyQt5图形界面

**新增功能**：
- **PyQt5 图形界面**（gui_app.py）：双击 exe 直接弹出界面
- **exe 打包**：两种打包形态（GUI版 + 命令行版）
- **延迟导入机制**：解决 PyQt5 与 win32com 的 COM 初始化冲突

---

### v3.8 (2026-04-28) - Step4全部用COM完成，解决I列超范围问题

**问题修复**：
- 代码生成的汇总表 I 列超出理论范围的问题
- 原因：openpyxl 赋值不触发 RANDBETWEEN 重算
- 解决：Step4 全部改用 COM 完成，模拟手工操作触发 F 列重新计算

---
### v3.7 (2026-04-22) - 代码质量优化 + 配置区重组

**共享模块**：`utils.py` 抽取 IDE 检测 + 颜色函数，消除两个文件间的重复代码

**重构**：
- `with_retry` 装饰器：天气 API 三个函数减少约 45 行重复代码
- `main()` 拆分为 `parse_user_input()` + `process_single_batch()` + `main()`
- 配置区重组：顶部「使用者配置区」集中所有可修改配置

**Bug修复**：`process_excel` 中 `wb_data` 未定义 bug

---

### v3.6 (2026-04-22) - WPS 兼容性修复

**问题**：WPS 环境下 `Visible=False` 时 pywinauto 无法点击"更新"对话框
**解决**：新增 `EXCEL_TYPE` + `EXCEL_VISIBLE` 配置，WPS 自动用 `True`，Office 用 `False`
**修改**：只需改顶部 `EXCEL_TYPE` 一行

---

### v3.5 (2026-04-20) - Excel静默运行 + 项目整理

**Excel 无窗口运行**：
- `excel.Visible = False`：完全隐藏 Excel 窗口
- `excel.ScreenUpdating = False`：关闭屏幕更新，提升性能
- `excel.DisplayAlerts = False`：禁止弹窗
- pywinauto 仍会自动点击隐藏的对话框按钮

**项目结构整理**：
- 根目录重命名：`project` → `rpt-maker`
- `dispatcher.py` → `run_queue.py`
- `archive/` → `.archive/`
- `logs/` → `.logs/`
- `待处理队列/` → `queue/`
- `已完成/` → `done/`
- 项目迁移至：`F:\myproject\rpt-maker`

---

### v3.4 (2026-04-17) - 自动化断点 + 健壮性增强 + 用户体验优化

**run_queue.py 新增功能**：
- `--notify / -n`：处理完成后发 Windows Toast 通知
- `--retry-failed`：读取 `.logs/failed_projects.txt`，只重试历史失败的工程
- `--dry-run`：只显示要做什么，不实际执行
- 失败记录持久化：成功后自动移除，失败时加入

**report_maker.py 健壮性增强**：
- Step5 文件级重试：每个 Excel 文件最多重试 3 次，指数退避
- PDF 转换重试：PDF 生成失败自动重试最多 3 次
- 幂等性保护：`copy_directory` 加 `force` 参数
- 天气 API 指数退避：三个 API 各最多重试 3 次

**report_maker.py 用户体验**：
- Windows Console 进度条：绕过 stdout 管道缓冲，TTY 下实时更新
- 成功/失败颜色：绿色/红色/青色/黄色
- 可选 tqdm 进度条：已安装时可用，未安装自动退化

**run_queue.py 输出优化**：
- 实时逐行输出：子进程输出不再延迟
- 编码修复：彻底解决跨编码环境乱码问题

---

### v3.3 (2026-04-16) - 封面修改逻辑重构

- 封面搜索范围优化：min_col=5, max_col=40
- 支持数值类型期数单元格：`str(cell.value)` 统一处理
- 天气前缀智能处理：保留"天气："前缀
- 去掉温度相关所有代码
- 版本升至 v3.3

---

### v3.2 (2026-04-16) - 多工程调度器 + 三重天气系统

- 多工程调度器 `run_all_projects.py`（后改名为 run_queue.py）
- 三重天气API系统（心知 + Open-Meteo Forecast + Historical）
- 环境变量支持：`REPORT_MAKER_PARENT`

---

### v3.1 (2026-04-15) - Step5 全自动化

- **核心突破**：pywinauto + COM 混合方案成功
- 7个纯 COM 方案失败后找到根本原因
- 13项代码优化完成

---

### v3.0 (2026-04-14) - 批量处理框架完成

- 六步主流程全部实现
- 7个 Step5 版本全部失败，找到根本原因

---

### v2.0 (2026-04-13) - 单期版完善 + 天气API集成

- 天气 API 集成
- 关键词列表遍历匹配

---

### v1.0 (2026-04-08) - 初始版本

- 基础功能实现

</details>

<details>
<summary><strong>📌 Step5自动化历程 — 点击展开</strong></summary>

---

## 📌 Step5自动化历程（完整记录）

### 7个失败版本（2026-04-14）

| 版本 | 方案 | 结果 |
|------|------|------|
| v1 | 纯COM，taskkill在CoInitialize后，无sleep | crash 0xC0000005 |
| v2 | taskkill移到CoInitialize之前 | 还是 crash |
| v3 | 拆成两个函数 | 变量没更新，AttributeError |
| v4 | 修复实例传回问题 | COM单例，第2批报"Open方法无效" |
| v5 | Quit后加sleep(2) | 仍报Open方法无效 |
| v6 | 注册表改UpdateLinks | 朱无理拒绝 |
| v7 | 重新结构化 | 第2批仍然失败 |

**根本原因**：Windows COM单例机制，Dispatch()永远返回已损坏的残存实例

### 最终成功方案

**pywinauto + COM 混合**：
- COM 负责打开/保存
- pywinauto 在 GUI 层点击"更新(&U)"按钮
- 不依赖 COM 状态

**v3.5 升级**：Excel 窗口隐藏运行，pywinauto 仍能点击隐藏的对话框

---

## 📌 问题与解决（完整记录）

| # | 问题 | 发现时间 | 解决 |
|---|------|---------|------|
| 1 | openpyxl写公式不写值 | 2026-04-10 | data_only=True补救 |
| 2 | COM冲突（openpyxl vs xlwings）| 2026-04-11 | 只用一个库 |
| 3 | 天气关键词匹配问题 | 2026-04-13 | 关键词列表遍历 |
| 4 | 天气和温度是分开的单元格 | 2026-04-10 | 正则r'\d+°C' |
| 5 | 日期格的两种表示 | 2026-04-10 | is_date_cell()兼容 |
| 6 | get_year_from_excel找不到日期 | 2026-04-14 | 兼容datetime和字符串 |
| 7 | da3te typo | 2026-04-14 | datetime.now().date() |
| 8 | 汇总表.xlsx错误提示 | - | 不影响功能，无需处理 |
| 9 | wb_data资源泄漏 | 2026-04-15 | try/finally保证close() |
| 10 | wb资源泄漏 | 2026-04-15 | try/finally保证close() |
| 11 | IndentationError in test3.py | 2026-04-13 | 添加pass或实际逻辑 |
| 12 | 天气API days参数从0开始 | 2026-04-09/10 | 改为days_diff+1 |
| 13 | 温度季节配置写错 | 2026-04-10 | 修正月份范围(3,4,5) |

---

## 📌 代码优化记录（完整）

### 13项代码优化（2026-04-15）
1. 删掉 `_close_activation_wizard()`（死代码）
2. 简化 `open_and_update_links` 的 except 块
3. 清理 `renamed` 变量（冗余）
4. 简化 main 循环中间变量
5. 去掉 winreg 和 pywinauto 的 try/except 包裹
6. 删掉 main() 里的 if winreg is None 检查（死代码）
7. `open_and_update_links` 重试逻辑重构
8. `_click_excel_update_button` 的 Desktop 对象缓存
9. `fetch_api_weather` 去掉 abs()
10. `get_year_from_excel` 加 try/finally
11. `process_excel` 里 wb_data 加 try/finally
12. 输入格式从多次 input 改成一次性逗号分隔
13. 调试打印加到 _do() 里（已注释备用）

### 4项后续优化（2026-04-16 上午）
14. 提取共享常量：WMO_CODE_MAP 和 JINAN_LAT/LON
15. 清理顶部 import
16. 简化资源管理
17. 稳定增强：`excel_to_pdf` 去掉多余的杀进程调用

### 封面修改优化（2026-04-16 下午）
18. 搜索范围改为 min_col=5, max_col=40
19. str() 统一处理，兼容数值类型
20. 天气前缀智能处理
21. 去掉温度相关所有代码

### v3.5 优化（2026-04-20）
22. Excel 无窗口运行：`Visible=False` + `ScreenUpdating=False` + `DisplayAlerts=False`
23. 项目结构整理：重命名 + 迁移路径

---

## 📁 历史文件对应

| 原始文件名 | 新文件名 | 版本 | 状态 |
|-----------|----------|------|------|
| test.py | report_maker.single.py | v1.0 | 存档 |
| run_all.py | report_maker.manual.py | v2.0 | 存档 |
| run_all1.py | report_maker.py | v3.0+ | 当前 |
| run_all1_backup.py | report_maker.backup.py | v3.1 | 存档 |
| run_all_projects.py | run_queue.py | v3.2 | 当前 |

</details>

<details>
<summary><strong>📞 技术支持 — 点击展开</strong></summary>

---

## 📞 技术支持

### 问题排查步骤
1. 查看控制台输出
2. 检查日志文件：`F:\myproject\rpt-maker\.logs\`
3. 启用调试打印
4. 检查Excel文件状态

### 关键文件位置
- **项目目录**: `F:\myproject\rpt-maker`
- **队列目录**: `F:\myproject\rpt-maker\queue`
- **已完成目录**: `F:\myproject\rpt-maker\queue\done`
- **日志目录**: `F:\myproject\rpt-maker\.logs`

### 开发记录
- `PROJECT_LOG.txt`（详细开发日志，逐日记录）
- `CHANGELOG.md`（版本变更摘要）

---

## 📊 项目状态

| 项目 | 内容 |
|------|------|
| 当前版本 | **v3.11.3** |
| 最后更新 | 2026-07-22 |
| 维护者 | 朱无理 |
| 项目历时 | 2026-04-08 至 2026-07-22 |
| Python版本 | 3.8+ |
| 操作系统 | Windows 10/11 |
| 依赖库 | openpyxl, pywin32, pywinauto, requests, PyQt5, pytz |

---

*最后更新：2026-07-22*
*版本：v3.11.3*

</details>
