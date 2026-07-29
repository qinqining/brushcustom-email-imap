# brushcustom 邮件询盘项目 checkpoint

更新日期：2026-07-28

## 项目目的

本项目用于只读读取企业微信/腾讯企业邮箱里的询盘邮件，筛选有效产品询盘，并整理到 Excel 表格中，便于 brushcustom 后续跟进客户、产品规格、公司信息、地理位置、附件资料和网站表单图片。

## 项目目录

项目位置：

```text
D:\brushcustom-email-imap
```

当前主 Excel 文件：

```text
D:\brushcustom-email-imap\brushcustom询盘记录_updated.xlsx
```

附件根目录：

```text
D:\brushcustom-email-imap\attachments
```

抓取和中间文件目录：

```text
D:\brushcustom-email-imap\work
```

重要文件：

```text
D:\brushcustom-email-imap\.env
D:\brushcustom-email-imap\fetch_imap.py
D:\brushcustom-email-imap\run_fetch_readonly.ps1
D:\brushcustom-email-imap\search_mail_readonly.py
D:\brushcustom-email-imap\save_mail_attachments_readonly.py
D:\brushcustom-email-imap\organize_historical_attachments.py
D:\brushcustom-email-imap\add_attachment_columns_20260724.mjs
D:\brushcustom-email-imap\checkpoint.md
```

`.env` 保存邮箱地址、IMAP 授权码、服务器等配置。不要把 `.env`、授权码或邮箱密码发给别人，也不要写进聊天、checkpoint 或公开仓库。

## 邮件读取安全规则

硬规则：绝不修改、移动、删除邮件，只做读取。

当前读取脚本必须遵守：

- 使用 `select(readonly=True)` 只读打开邮箱。
- 使用 `BODY.PEEK[]` 获取邮件内容，避免把邮件标记为已读。
- 禁止 `STORE`、`EXPUNGE`、`CLOSE`、`DELETE`、`MOVE`、`COPY`、`APPEND`、`CREATE`、`RENAME`、`SUBSCRIBE`、`UNSUBSCRIBE`、`UID STORE` 等可能改变邮箱状态的 IMAP 操作。
- 默认读取 `MAILBOX=INBOX`。

固定读取入口示例：

```powershell
cd /d D:\brushcustom-email-imap
.\run_fetch_readonly.ps1 -Start 2026-07-10
```

读取日期区间时，`--start` 包含当天，`--end` 不包含当天。例如读取 2026-07-07 到 2026-07-10：

```powershell
.\run_fetch_readonly.ps1 -Start 2026-07-07 -End 2026-07-11
```

## Excel 字段

当前主表有 13 列：

```text
A 日期
B 发件人姓名
C 发件邮箱
D 咨询产品及详细规格
E 客户公司名称
F 公司官方网站
G 所属行业
H 公司规模
I 年营收估计
J 客户地理位置
K 状态/备注
L 附件/资料路径
M 原始图片URL
```

附件整理规则：

- 不把缩略图直接嵌入 Excel。
- 邮件附件、网站表单上传图片、图纸、STEP 文件等统一保存到 `D:\brushcustom-email-imap\attachments` 下的客户文件夹。
- Excel 的 `附件/资料路径` 列填写对应客户附件文件夹路径。
- Excel 的 `原始图片URL` 列只填写网站表单上传图片的原始 URL；普通邮件附件没有原始 URL 时留空。
- 每个附件文件夹可包含 `inquiry_email.txt`、`matched_mailbox.txt`、`original_urls.txt`、`notes.txt` 等辅助记录。

## 询盘筛选和录入规则

录入：

- 产品询盘
- RFQ / 报价请求
- 样品请求
- 规格确认
- 客户主动提供产品尺寸、数量、用途、图纸、附件或采购需求的邮件

不录入：

- 企业微信登录提醒、安全提醒
- 欺诈/付款异常沟通本身，除非客户明确提出继续采购或恢复合作
- 纯广告、推广、SEO、友链、建站服务
- 无具体产品、规格、数量、应用场景的泛询价
- 已经存在于 Excel 的重复客户邮件，除非是新的正式规格补充或新需求

公司信息规则：

- 官网、行业、规模、营收、地理位置能确认就填写。
- 查不到或不确定的字段统一写 `待确认`。
- 不要为了填满表格而猜测公司规模或营收。

## 当前进度

主表当前有 16 条询盘/线索数据行。

当前主表已含：

1. Emre Makina Sanayi
2. TNT Group
3. Specialist Metallic Coatings
4. Volkswagen (VW)
5. Lien Minh Industrial Equipment Joint Stock Company
6. Carsoe
7. ADVANX TECHNOLOGY PTE LTD
8. ACN Cutting Systems (Motofil Group)
9. CISA, CEDACERIA INDUSTRIAL, S.L.
10. Noventa (Thailand) Co., Ltd.
11. Nautica Technologies AG 访问请求
12. AssetCool / Cable Coatings Limited
13. Nichiha USA, Inc.
14. American Flexible Products
15. Nautica Technologies AG 正式产品规格补充
16. Rhino Roofing Products Ltd.

2026-07-24 已完成历史附件重新整理：

- 已新增 `状态/备注`、`附件/资料路径`、`原始图片URL` 三列。
- CISA 行原有状态 `已下单` 已保留。
- 已为 10 条有附件或网站上传图片的历史询盘填写附件路径。
- 已为 Nichiha 和 American Flexible Products 填写网站表单上传图片的原始 URL。
- 未嵌入缩略图。

已整理到附件目录的记录：

```text
D:\brushcustom-email-imap\attachments\2026-06-16_Emre-Makina_Mehmet-Okan-Yerli
D:\brushcustom-email-imap\attachments\2026-06-18_TNT-Group_Gigi-Lam
D:\brushcustom-email-imap\attachments\2026-06-26_Lien-Minh_Trang
D:\brushcustom-email-imap\attachments\2026-07-09_ADVANX_Rena-Goh
D:\brushcustom-email-imap\attachments\2026-07-10_ACN-Cutting_Maria-Clara-Cabrita
D:\brushcustom-email-imap\attachments\2026-07-14_CISA_Rita-Finale
D:\brushcustom-email-imap\attachments\2026-07-16_Noventa_Kodchaphan
D:\brushcustom-email-imap\attachments\2026-07-17_AssetCool_Jasper-Poppele
D:\brushcustom-email-imap\attachments\2026-07-20_Nichiha_Jim-Farina
D:\brushcustom-email-imap\attachments\2026-07-23_American-Flexible_Zory
D:\brushcustom-email-imap\attachments\2026-07-27_Rhino-Roofing-Products_Wasantha
```

特别说明：

- Emre Makina 和 TNT Group / Gigi Lam 的历史邮件在当前可搜索邮箱中未能重新定位到原邮件，所以对应附件文件夹目前只有说明文件，实际附件待后续如能恢复原邮件再补。
- Nichiha 原始图片 URL：`https://brushcustom.com/wp-content/uploads/elementor/forms/6a5e2d7a5e430.jpg`
- American Flexible Products 原始图片 URL：`https://brushcustom.com/wp-content/uploads/elementor/forms/6a6215f58e30f.jpg`
- Specialist Metallic Coatings、Volkswagen、Carsoe、Nautica 两条记录当前没有发现需要保存的附件或原始图片 URL。
- Rhino Roofing Products 的 `image002.jpg` 是邮件签名/logo，不是产品图纸；已保存但报价仍需客户提供技术图纸或由我方根据规格出图。

2026-07-28 新增 Rhino 询盘整理：

- 邮件日期：2026-07-27 10:50:40 UTC。
- 发件人：`imports@rhino.lk`，联系人署名 Wasantha。
- 产品：Upper Cleaning Brush Fibre fitted。
- 数量：02 Nos。
- 规格：`ø 350 / 254.75 x 1200 mm`。
- 报价要求：C&F Colombo，需提供品牌/原产国、技术数据表、图纸、交期、质保期、付款条款。
- 公司：Rhino Roofing Products Ltd.，官网 `https://www.rhino.lk/`，地址为 Sri Lanka, Colombo 09。
- 年营收写 `待确认`。
- 已保存附件目录：`D:\brushcustom-email-imap\attachments\2026-07-27_Rhino-Roofing-Products_Wasantha`。
- Rhino 已正式写入主表第 17 行。
- 当时生成过 pending 更新文件，用户关闭 Excel 后已覆盖回主表，pending 文件和 Rhino 写入前备份文件已删除。
- 主表校验结果：

```text
officecli validate: Validation passed: no errors found.
officecli view issues: Found 0 issue(s).
```

- 预览图：

```text
D:\brushcustom-email-imap\work\preview_after_rhino_20260727.png
```

本次写表备份：

```text
D:\brushcustom-email-imap\work\brushcustom询盘记录_backup_before_attachment_columns_20260724022013.xlsx
```

本次 Excel 校验结果：

```text
officecli validate: Validation passed: no errors found.
officecli view issues: Found 0 issue(s).
```

预览图：

```text
D:\brushcustom-email-imap\work\preview_after_attachment_columns_20260724.png
```

## 后续常用操作

只读读取近期邮件：

```powershell
cd /d D:\brushcustom-email-imap
.\run_fetch_readonly.ps1 -Start 2026-07-24
```

保存邮件附件时使用：

```powershell
python .\save_mail_attachments_readonly.py
```

历史附件整理脚本：

```powershell
python .\organize_historical_attachments.py
```

本次给 Excel 补附件列的脚本：

```powershell
node .\add_attachment_columns_20260724.mjs
```

校验 Excel：

```powershell
officecli validate "D:\brushcustom-email-imap\brushcustom询盘记录_updated.xlsx"
officecli view "D:\brushcustom-email-imap\brushcustom询盘记录_updated.xlsx" issues
```

## 换设备恢复步骤

1. 复制整个项目目录到新设备：

```text
D:\brushcustom-email-imap
```

2. 确认 `.env` 存在，并填写新设备可用的邮箱地址和 IMAP 客户端授权码。
3. 确认主 Excel 文件仍在：

```text
D:\brushcustom-email-imap\brushcustom询盘记录_updated.xlsx
```

4. 确认 OfficeCLI skills 已安装：

```text
C:\Users\HP\.codex\skills\officecli
C:\Users\HP\.codex\skills\officecli-xlsx
```

5. 先用只读命令验证邮箱读取，再筛选询盘、保存附件、写入 Excel。

## 注意事项

- 每次写 Excel 前先备份当前主表。
- 后续新询盘如果有附件，附件保存到 `attachments\日期_公司_联系人`，Excel 只写附件文件夹路径。
- 网站表单上传图片要同时记录本地保存路径和原始图片 URL。
- 如果 IMAP 授权码失效，需要在企业微信/腾讯企业邮箱重新生成客户端专用密码。
- 如果查不到公司官网、营收或规模，写 `待确认`。
