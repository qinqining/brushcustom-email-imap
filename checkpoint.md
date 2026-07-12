# brushcustom 邮件询盘项目 checkpoint

更新日期：2026-07-12

## 项目目的

本项目用于只读读取企业微信/腾讯企业邮箱收件箱中的询盘邮件，筛选有效产品询盘，并整理到 Excel 表格中，便于 brushcustom 后续跟进客户、产品规格、公司信息和地理位置。

## 项目目录

项目位置：

```text
D:\brushcustom-email-imap
```

当前主 Excel 文件：

```text
D:\brushcustom-email-imap\brushcustom询盘记录_updated.xlsx
```

重要文件：

```text
D:\brushcustom-email-imap\.env
D:\brushcustom-email-imap\fetch_imap.py
D:\brushcustom-email-imap\run_fetch_readonly.ps1
D:\brushcustom-email-imap\append_inquiries.mjs
D:\brushcustom-email-imap\format_final_workbook.mjs
D:\brushcustom-email-imap\README.md
D:\brushcustom-email-imap\checkpoint.md
```

`.env` 用于保存邮箱地址、IMAP 授权码、IMAP 服务器和 Excel 路径。不要把 `.env` 发给别人，也不要把授权码写进聊天或 checkpoint。

## 邮件读取规则

硬规则：绝不修改、移动、删除邮件，只做读取。

当前读取脚本已经做了只读护栏：

- 使用 `select(readonly=True)` 只读打开收件箱。
- 使用 `BODY.PEEK[]` 获取邮件内容，避免把邮件标记为已读。
- 禁用 `STORE`、`EXPUNGE`、`CLOSE`、`DELETE`、`MOVE`、`COPY`、`APPEND`、`CREATE`、`RENAME`、`SUBSCRIBE`、`UNSUBSCRIBE`、`UID STORE` 等可能改变邮箱状态的 IMAP 操作。
- 默认读取 `MAILBOX=INBOX`。

固定读取入口：

```powershell
.\run_fetch_readonly.ps1 -Start 2026-07-10
```

读取日期区间时，`--start` 包含当天，`--end` 不包含当天。例如读取 2026-07-07 到 2026-07-10：

```powershell
.\run_fetch_readonly.ps1 -Start 2026-07-07 -End 2026-07-11
```

抓取结果保存到：

```text
D:\brushcustom-email-imap\work
```

## 询盘筛选和录入规则

录入：

- 产品询盘
- RFQ / 报价请求
- 样品请求
- 规格确认
- 客户主动提供产品尺寸、数量、用途、图纸或采购需求的邮件

不录入：

- 企业微信登录提醒、安全提醒
- 欺诈/付款异常沟通
- 纯售后或非产品询盘
- 已经存在于 Excel 的重复客户邮件

去重规则：

- 优先按发件邮箱去重。
- 同一客户后续补充规格时，可以合并到原记录或新增备注，避免重复创建客户。

公司信息规则：

- 官网、行业、规模、营收、地理位置能确认就填写。
- 查不到或不确定的字段统一写 `待确认`。
- 不要为了填满表格而猜测公司规模或营收。

## 当前进度

已完成：

- 已在 `D:\brushcustom-email-imap\.env` 配置企业邮箱 IMAP 参数。
- 已成功只读连接 IMAP。
- 已验证读取 2026-07-10 收件箱邮件，抓到 11 封。
- 已筛掉企业微信安全提醒和非询盘邮件。
- 已把有效询盘整理进 Excel。
- 已安装 OfficeCLI skills：
  - `C:\Users\HP\.codex\skills\officecli`
  - `C:\Users\HP\.codex\skills\officecli-xlsx`
- 已验证 `officecli --version` 为 `1.0.129`。

当前 Excel 已有 8 条询盘数据行：

1. Carsoe
2. ADVANX TECHNOLOGY PTE LTD
3. ACN Cutting Systems (Motofil Group)
4. Emre Makina Sanayi
5. TNT Group
6. Specialist Metallic Coatings
7. Volkswagen (VW)
8. Lien Minh Industrial Equipment Joint Stock Company

2026-07-10 这次新增录入了 5 条：

- Emre Makina Sanayi
- TNT Group
- Specialist Metallic Coatings
- Volkswagen (VW)
- Lien Minh Industrial Equipment Joint Stock Company

## 换设备后的恢复步骤

1. 复制整个项目目录到新设备：

```text
D:\brushcustom-email-imap
```

2. 确认 `.env` 存在，并填写新设备可用的邮箱地址和 IMAP 客户端授权码。

3. 确认 Excel 文件路径仍是：

```text
D:\brushcustom-email-imap\brushcustom询盘记录_updated.xlsx
```

4. 安装或恢复 OfficeCLI skills 后，重启 Codex 让新 skills 生效。

5. 验证只读邮件读取：

```powershell
cd /d D:\brushcustom-email-imap
.\run_fetch_readonly.ps1 -Start 2026-07-10
```

6. 后续需要更新 Excel 时，先读取邮件 JSON，再筛选询盘、补充公司信息，最后写入主 Excel。所有未知公司字段写 `待确认`。

## 注意事项

- 不要把 `.env`、授权码、邮箱密码提交到公开仓库或发给他人。
- 每次写 Excel 前建议先备份当前主表。
- 如果 IMAP 授权码失效，需要在企业微信/腾讯企业邮箱重新生成客户端专用密码。
- 如果后续 Codex 重启后能识别 OfficeCLI skills，编辑 `.xlsx` 时优先使用 OfficeCLI；如工具不可用，可以继续使用当前项目已有脚本和表格工具。
