# DHL Retoure Automation

从 Excel 读取退货数据，在 DHL Geschäftskundenportal 中批量创建 `DHL Retoure Online` 标签，并把结果页 `Sendungsnummer` 写回 Excel 的 L 列。

## 重要提示

- 本工具会创建真实 DHL 退货委托，请先使用 `--dry-run` 检查第一行。
- 用户名、密码和登录 Cookie 不会写入项目或 Excel。
- 登录会话保存在当前电脑用户目录的 `.dhl-retoure-automation/chrome-profile` 中。
- 已有 `Sendungsnummer` 的行会自动跳过，支持中断后继续。
- 每成功创建一单就立即保存 Excel，减少中断导致的数据丢失。
- DHL 页面结构变化后，选择器可能需要更新。

## 环境要求

- Windows 10/11（当前版本主要在 Windows + Chrome 下使用）
- [Node.js 20 或更新版本](https://nodejs.org/)
- Google Chrome
- 有权限使用 DHL Geschäftskundenportal 的账户

## Excel 表头

第一张工作表第 1 行必须包含以下标题：

| Excel 标题 | DHL 字段 |
|---|---|
| Sendungsreferenz | Sendungsreferenz |
| KUNDENREFERENZ / AUF LABEL ANZEIGEN | Kundenreferenz |
| VOR-UND NACHNAME | Vor- und Nachname |
| NAMENSZUSATZ 1 | Namenszusatz 1 |
| NAMENSZUSATZ 2 | Namenszusatz 2 |
| PLZ | PLZ（四位邮编自动补 0） |
| ORT | Ort |
| STRASSE | Straße |
| NR | Hausnummer |
| E-MAIL ADRESSE DES KUNDEN | Kunden-E-Mail |
| Retouren-Empfänger | DHL 账户中的收件方名称 |
| Sendungsnummer | 输出列；没有时自动创建为 L 列 |

## 快速开始（Windows）

1. 下载并解压项目。
2. 双击 `start.bat`。第一次运行会自动安装依赖。
3. 将 Excel 文件拖进命令窗口并按 Enter。
4. Chrome 打开后自行登录 DHL。
5. 确认待处理数量，输入 `CREATE` 才会开始真实批量创建。
6. 输出文件默认保存在源 Excel 同目录，文件名增加 `_已填写运单号`。

## 推荐：先试运行

```powershell
npm install
npm start -- --input "C:\data\DHL退货标签.xlsx" --dry-run
```

试运行只填写第一条待处理记录，不点击 `Retoure beauftragen`。

确认无误后运行：

```powershell
npm start -- --input "C:\data\DHL退货标签.xlsx"
```

可选参数：

- `--output "路径.xlsx"`：指定输出文件。
- `--overwrite`：直接覆盖输入文件（默认不会覆盖）。
- `--dry-run`：只填第一条并停在提交前。
- `--headless`：无界面模式；首次登录不建议使用。

## 隐私与安全

请勿把真实客户 Excel、浏览器登录目录、日志或密码提交到 GitHub。项目的 `.gitignore` 已忽略 Excel、输出目录和本地浏览器会话。

## 免责声明

这是非官方工具，与 DHL 无隶属关系。使用者应确保自动化操作符合其 DHL 合同、内部流程和适用法律。建议先用少量数据测试，并人工抽查结果。

