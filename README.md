# brushcustom 邮件询盘读取项目

## 邮件读取硬规则

1. 只读取邮件，不修改、不移动、不删除邮件。
2. IMAP 必须使用只读选择邮箱：`select(readonly=True)`。
3. 抓取正文必须使用 `BODY.PEEK[]`，避免把邮件标记为已读。
4. 禁止使用这些 IMAP 操作：`STORE`、`EXPUNGE`、`CLOSE`、`DELETE`、`MOVE`、`COPY`、`APPEND`、`CREATE`、`RENAME`、`SUBSCRIBE`、`UNSUBSCRIBE`、`UID STORE`。
5. `.env` 里的邮箱地址和授权码不打印、不写入输出表。

## 询盘整理规则

1. 只整理产品询盘、报价请求、样品请求、规格确认等销售线索。
2. 企业微信登录提醒、安全提醒、付款欺诈沟通、纯售后或非询盘邮件不录入。
3. 已在 Excel 中存在的发件邮箱不重复录入。
4. 客户公司官网、公司规模、年营收等查不到或不确定时，一律写 `待确认`。
5. 日期范围使用明确日期：`--start` 为包含当天，`--end` 为不包含当天。

## 常用命令

读取 2026-07-10 当天邮件：

```powershell
.\run_fetch_readonly.ps1 -Start 2026-07-10
```

读取 2026-07-07 到 2026-07-10：

```powershell
.\run_fetch_readonly.ps1 -Start 2026-07-07 -End 2026-07-11
```

抓取结果会保存到 `work/` 目录下的 JSON 文件。后续需要编辑 Excel 时，可以使用已安装的 OfficeCLI skills，未知公司信息按 `待确认` 处理。
