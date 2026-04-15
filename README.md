# NapCat Notify

## Compare + Notify

`compare_and_notify.py` glues the two existing workflows together:

- export the latest Jinshuju entries
- compare them against the qualification Excel
- turn the chosen mismatch set into QQ recipients
- run a NapCat precheck and dry-run by default
- only send when you add `--send`

Common commands:

```powershell
python .\compare_and_notify.py `
  --entries-url "https://jinshuju.net/forms/SnQ2YZ/entries" `
  --qualification-file "C:\Users\Theta\Downloads\成电杯正赛资格名单.xlsx" `
  --profile-dir "C:\Users\Theta\napcat_notify\browser_state\jinshuju_profile_qsqni76m" `
  --group-id 1042339991 `
  --headless
```

```powershell
python .\compare_and_notify.py `
  --entries-url "https://jinshuju.net/forms/SnQ2YZ/entries" `
  --qualification-file "C:\Users\Theta\Downloads\成电杯正赛资格名单.xlsx" `
  --profile-dir "C:\Users\Theta\napcat_notify\browser_state\jinshuju_profile_qsqni76m" `
  --group-id 1042339991 `
  --headless `
  --send
```

Double-click launcher:

```text
start_compare_and_notify.bat
```

Extra outputs:

- `notify_candidates.csv`
- `notify_skipped.csv`
- `notify\results.csv`
- `notify\results.jsonl`
- `notify\notify_summary.json`

用于两类本地自动化任务：

- 通过 NapCat OneBot 11 批量发送 QQ 私聊或群临时会话通知
- 通过浏览器自动化从金数据后台导出表单，再和本地资格名单比对

当前项目包含两个主要脚本：

- `notify.py`：批量 QQ 通知
- `compare_jinshuju.py`：金数据报名比对
- `start_notify.bat`：一键启动 QQ 通知
- `start_compare_jinshuju.bat`：一键启动金数据比对

## 环境要求

- Windows
- Python 3.10+
- 已安装依赖：`pip install -r requirements.txt`
- 如果要运行金数据比对脚本，还需要额外安装浏览器内核：

```powershell
playwright install chromium
```

## 快速启动

```powershell
git clone https://github.com/nesettle/napcat_notify.git
cd napcat_notify
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
playwright install chromium
```

如果 PowerShell 默认禁止脚本执行，可以先运行：

```powershell
Set-ExecutionPolicy -Scope Process Bypass
```

## QQ 通知脚本

`notify.py` 用于通过 NapCat OneBot 11 向固定名单发送提醒消息。

### NapCat 前置条件

- QQ 已登录
- NapCat 已启动
- OneBot 11 `WS 服务端` 已启用
- 默认监听地址可用：`ws://127.0.0.1:3001`

默认读取的 NapCat 配置文件路径：

```text
C:\ProgramData\NapCatQQ Desktop\runtime\NapCatQQ\config\onebot11_3496930386.json
```

### 常用命令

```powershell
python .\notify.py
python .\notify.py --send --limit 1
python .\notify.py --send --group-id 1042339991
python .\notify.py --input C:\path\to\recipients.xlsx
```

也可以直接双击：

```text
start_notify.bat
```

### 支持能力

- 默认 dry-run
- 内置 21 人名单
- 支持 `CSV/XLSX` 名单导入
- 支持普通私聊发送
- 支持群临时会话发送
- 自动输出 `results.csv` 与 `results.jsonl`

## 金数据报名比对脚本

`compare_jinshuju.py` 会打开金数据后台首页，定位指定表单，进入数据页，导出报名数据，再与本地资格名单对比，输出：

- `qualified_not_registered.csv`
- `registered_not_qualified.csv`
- `matched.csv`
- `qualification_duplicates.csv`
- `form_duplicates.csv`
- `summary.json`

如果你已经知道数据页直达链接，也可以跳过首页找表单，直接传 `--entries-url`。

### 工作方式

- 使用持久化浏览器目录保存登录态
- 第一次运行时，如果未登录金数据，会打开浏览器让你手动登录
- 登录完成后，脚本会在首页按表单标题寻找对应表单
- 进入数据页后，自动执行“更多 -> 导出数据 -> 下载”
- 下载完成后，解析导出文件并与本地资格名单比对

### 常用命令

```powershell
python .\compare_jinshuju.py `
  --form-title "成电杯正赛单打项目报名" `
  --qualification-file "C:\Users\Theta\Downloads\成电杯正赛资格名单.xlsx"
```

直接使用数据页 URL：

```powershell
python .\compare_jinshuju.py `
  --entries-url "https://jinshuju.net/forms/SnQ2YZ/entries" `
  --qualification-file "C:\Users\Theta\Downloads\成电杯正赛资格名单.xlsx"
```

也可以直接双击：

```text
start_compare_jinshuju.bat
```

只统计某个时间之后的报名：

```powershell
python .\compare_jinshuju.py `
  --form-title "成电杯正赛单打项目报名" `
  --qualification-file "C:\Users\Theta\Downloads\成电杯正赛资格名单.xlsx" `
  --created-after 2026-04-14
```

如果导出表中的列名不是默认的 `姓名 / 学院 / QQ号`，可以覆盖：

```powershell
python .\compare_jinshuju.py `
  --form-title "成电杯正赛单打项目报名" `
  --qualification-file "C:\Users\Theta\Downloads\成电杯正赛资格名单.xlsx" `
  --name-field-label "姓名" `
  --college-field-label "学院/部门" `
  --qq-field-label "QQ号"
```

### 主要参数

| 参数 | 说明 |
| --- | --- |
| `--form-title` | 金数据后台首页中的表单标题 |
| `--entries-url` | 金数据数据页直达链接；提供后可跳过首页定位 |
| `--qualification-file` | 本地资格名单 Excel 文件 |
| `--name-field-label` | 导出表中的姓名列名，默认 `姓名` |
| `--college-field-label` | 导出表中的学院列名，默认 `学院` |
| `--qq-field-label` | 导出表中的 QQ 列名，默认 `QQ号` |
| `--created-after` | 只统计该时间之后创建的报名 |
| `--download-format` | 导出格式，支持 `xlsx` / `csv`，默认 `xlsx` |
| `--profile-dir` | 浏览器持久化目录，默认 `browser_state\jinshuju_profile` |
| `--output-root` | 输出根目录，默认 `runs` |
| `--headless` | 无头模式运行；首次登录不建议使用 |

### 匹配规则

- 优先按 QQ 号匹配
- 如果一侧 QQ 缺失，回退到 `姓名 + 学院`
- 姓名会自动去除尾部备注，例如 `王毓远（领队） -> 王毓远`
- 学院会统一空格与中英文括号
- 金数据重复提交时，主比较只保留最新一条

### 输出目录

每次运行都会在 `runs` 下创建时间戳目录，例如：

```text
C:\Users\Theta\napcat_notify\runs\compare-20260415-221530\
```

其中至少包含：

- `jinshuju_export.xlsx` 或 `jinshuju_export.csv`
- `matched.csv`
- `qualified_not_registered.csv`
- `registered_not_qualified.csv`
- `qualification_duplicates.csv`
- `form_duplicates.csv`
- `summary.json`

如果浏览器自动化中途失败，还会额外保存：

- `jinshuju_debug.html`
- `jinshuju_debug.png`

### 注意事项

- 这版不依赖金数据 API，因此不要求升级到支持 API 的订阅套餐
- 这版依赖金数据后台页面结构；如果页面按钮文案变化，可能需要调整脚本
- 首次登录必须使用可见浏览器窗口
- 如果你的表单标题在首页里不唯一，脚本会点击第一个命中的表单

## 当前验证情况

已验证通过：

- `notify.py` 的 dry-run、单条试发、群临时会话发送
- `compare_jinshuju.py` 的语法校验
- 资格名单去重与姓名规范化逻辑
- 按 QQ 优先、姓名+学院兜底的核心比对逻辑
