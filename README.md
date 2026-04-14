# NapCat Notify

基于 NapCat OneBot 11 `WS 服务端` 的批量私聊通知脚本，适合做名单通知、报名提醒和结果留档。

当前脚本默认面向“有资格但未报名单打”的名单，支持：

- 使用内置名单直接发送
- 从 `CSV/XLSX` 读取收件人
- `dry-run` 预览消息和对象
- 单条试发
- 全量发送
- 普通私聊发送
- 基于群号的群临时会话发送
- 自动输出 `CSV/JSONL` 结果日志

## 功能概览

- 启动时自动读取 NapCat 的 OneBot 11 配置
- 自动连接本机 `WS 服务端`
- 先做 `get_login_info` 预检，确认 bot 在线
- 支持按姓名个性化消息模板
- 发送结果自动记录为：
  - `sent`
  - `api_failed`
  - `transport_failed`
  - `dry_run`

## 目录结构

```text
napcat_notify/
├─ notify.py
├─ README.md
├─ recipients.sample.csv
├─ recipients.sample.xlsx
└─ runs/
   └─ <timestamp>/
      ├─ results.csv
      └─ results.jsonl
```

## 环境要求

- Windows
- Python 3.10+
- NapCat 已启动并可正常工作
- NapCat OneBot 11 `WS 服务端` 已配置完成

默认读取的配置文件路径：

```text
C:\ProgramData\NapCatQQ Desktop\runtime\NapCatQQ\config\onebot11_3496930386.json
```

## 默认行为

- 默认只执行 `dry-run`
- 默认使用内置的 21 人名单
- 默认消息间隔为 `2` 秒
- 默认不重试失败消息
- 默认不自动加好友
- 默认不做失败补救

## 默认消息模板

```text
{name}同学你好，我是乒协机器人，这边发现你在正赛资格名单中，但单打项目报名表中没有你的信息。请你确认是否参加成电杯单打项目，如果参加，请填写群里的单打报名链接。
```

脚本会自动把 `{name}` 替换成对应姓名。

## 用法

### 1. 只做 dry-run

```powershell
python C:\Users\Theta\napcat_notify\notify.py
```

### 2. 单条试发

```powershell
python C:\Users\Theta\napcat_notify\notify.py --send --limit 1
```

### 3. 全量发送

```powershell
python C:\Users\Theta\napcat_notify\notify.py --send
```

### 4. 使用群临时会话试发 1 人

```powershell
python C:\Users\Theta\napcat_notify\notify.py --send --limit 1 --group-id 1042339991
```

### 5. 使用群临时会话全量发送

```powershell
python C:\Users\Theta\napcat_notify\notify.py --send --group-id 1042339991
```

### 6. 从文件读取名单

```powershell
python C:\Users\Theta\napcat_notify\notify.py --input C:\path\to\recipients.csv
python C:\Users\Theta\napcat_notify\notify.py --input C:\path\to\recipients.xlsx
```

## 参数说明

| 参数 | 说明 |
| --- | --- |
| `--input` | 收件人文件路径，支持 `.csv` / `.xlsx` |
| `--config` | NapCat OneBot 配置文件路径 |
| `--output-root` | 结果输出目录根路径 |
| `--send` | 真正发送消息；不传时只做 `dry-run` |
| `--limit` | 只处理前 N 条记录 |
| `--delay` | 每条消息之间的间隔秒数 |
| `--group-id` | 启用群临时会话发送时使用的群号 |

## 输入文件格式

必需列：

- `QQ号`
- `姓名`
- `学院`

可选列：

- `消息`

规则：

- 如果提供 `消息` 列，则该行优先使用自定义消息
- 如果不提供 `消息`，则使用默认模板
- `QQ号` 会自动清洗为纯数字字符串

## 群临时会话说明

如果目标不是好友，但和 bot 在同一个群里，可以提供 `--group-id`。

脚本会先调用 `get_group_member_info` 校验该用户是否仍在该群中，校验通过后再使用带 `group_id` 的 `send_private_msg` 发消息。

如果用户不在该群中，会记录为 `api_failed`，不会继续对该对象发送。

## 输出结果

每次运行都会生成一个时间戳目录，例如：

```text
C:\Users\Theta\napcat_notify\runs\20260414-210831\
```

输出文件：

- `results.csv`
- `results.jsonl`

字段：

- `qq`
- `name`
- `college`
- `mode`
- `status`
- `message_id`
- `error`
- `sent_at`
- `message_preview`

## 常见问题

### 1. 提示“请先添加对方为好友”

说明普通私聊发不出去。优先改用群临时会话：

```powershell
python C:\Users\Theta\napcat_notify\notify.py --send --group-id <群号>
```

### 2. 连接预检失败

先确认：

- NapCat 已启动
- QQ 已登录
- `onebot11_3496930386.json` 中的 `websocketServers` 已启用
- 对应端口正在监听

### 3. 文件能读但中文乱码

`CSV` 建议保存为 `UTF-8 with BOM`。  
`XLSX` 推荐直接使用 Excel/WPS 保存。

## 当前验证情况

已验证通过：

- 默认 21 人 `dry-run`
- `CSV` 输入 `dry-run`
- `XLSX` 输入 `dry-run`
- 普通私聊单条试发
- 群临时会话单条试发
- 群临时会话全量发送

