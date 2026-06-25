# 附件检查助手 (Outlook Attachment Checker)

发送邮件时，如果正文提到了「附件 / 随函附送 / attached」等字样，但还没有添加附件，
就在发送前弹出提醒。所有检查都在本机 Outlook 内完成，邮件内容不会上传。

这是一个**发送时事件 (Smart Alerts / `OnMessageSend`)** 类型的 Outlook 加载项。

---

## 排查记录：“无法加载这个加载项 / 一直卡在正在处理”

通过在网页版 Outlook 用 F12 控制台抓取真实日志，逐项排除后定位到根因：

- 托管正常：文件部署在 GitHub Pages（`https://kuroi-sakuya.github.io/outlook-attachment-checker/`），
  HTTPS、MIME 类型、可达性都没问题。
- 账户正常：Microsoft 365 工作/学校邮箱（Exchange Online），支持发送时事件。
- 清单与代码结构正常。
- **真正的症状**：点发送后，控制台只出现 `>>> ACHK ...: 脚本已加载，... 已关联完成`，
  **再也没有 `onMessageSendHandler 已触发`**。即——**脚本加载并注册成功了，但点发送时
  处理函数根本没有被调用**，于是永远不会调用 `event.completed`，Outlook 就卡在
  “附件检查助手 正在处理邮件… 所花时间超过预期”。

“注册了却没被调用”最常见的两个原因：

1. 管理中心实际部署的清单与代码里的 `FunctionName` 对不上（叫了一个没注册的函数名）。
2. 线上文件 / 部署被缓存成了旧状态。

### 修复（v4，一次性覆盖以上两种）

- `launchevent.v4.js`：在 v3 基础上加固——
  1. **多函数名注册**：用多个常见候选名都关联到同一处理函数，无论部署的清单写的是哪个名字都能命中；
  2. **硬性兜底超时**（3.5 秒必放行）+ 每个异步调用单独的短超时，彻底杜绝“卡死”；
  3. 日志横幅改为 `>>> ACHK v4:`，方便确认新文件是否真正生效。
- `commands.v4.html`：引用 v4 脚本的运行载体页。
- `manifest.xml`：保留原 `<Id>`（在管理中心里是“更新”而非“新增”），运行时改指向
  **新文件名 v4**（强制绕开缓存），`FunctionName` 保持 `onMessageSendHandler`。

### 部署 & 验证步骤

1. 把本分支合并到 `main`（GitHub Pages 从 `main` 根目录发布，约 1–2 分钟后新文件上线）。
2. 浏览器打开确认新文件已上线、能看到内容：
   - https://kuroi-sakuya.github.io/outlook-attachment-checker/launchevent.v4.js
   - https://kuroi-sakuya.github.io/outlook-attachment-checker/commands.v4.html
3. 请管理员在 **Microsoft 365 管理中心 → 集成应用** 用本仓库的 `manifest.xml`
   **更新**该加载项（建议先移除旧的再重新部署，确保干净刷新）。
4. 在网页版 Outlook 里按 F12 → 控制台 Filter 输入 `ACHK` → 清空 → 写一封正文含“见附件”
   的邮件点发送，按下表判断：

| 控制台现象 | 含义 |
| --- | --- |
| 看到 `ACHK v4 … 已触发` 且能正常拦截/放行 | ✅ 修复成功 |
| 看到 `ACHK v4` 横幅、但仍只停在“脚本已加载”、无“已触发” | 部署的清单没指到本函数 → 让管理员**移除后重新部署**本 `manifest.xml` |
| 连 `ACHK v4` 横幅都没有 | 线上还是旧文件 → 确认已合并到 `main` 且 Pages 已发布、清单已更新 |

---

## 技术要点

- 发送时事件要求 **Mailbox 需求集 1.12**，且邮箱须为 **Exchange Online**（个人 Outlook.com、
  Gmail/IMAP、本地 Exchange 均不支持）。
- 处理函数通过 `Office.actions.associate("onMessageSendHandler", ...)` 注册；
  清单 `<LaunchEvent FunctionName>` 必须与之一致。
- `SendMode="PromptUser"`：提醒后用户仍可选择“仍然发送”。
- 托管主机要求：HTTPS + 正确 MIME 类型；GitHub Pages 满足，
  但 `raw.githubusercontent.com`、`github.com/.../blob/...` **不行**（会被当成纯文本/网页外壳）。

## 文件说明

| 文件 | 作用 |
| --- | --- |
| `manifest.xml` | 加载项清单（提交给管理中心的就是它），当前指向 v4 |
| `commands.v4.html` / `launchevent.v4.js` | **当前使用**：运行载体页 + 加固后的发送时检查逻辑 |
| `taskpane.html` | 功能区按钮打开的信息展示页 |
| `assets/` | 图标 |
| `*.v3.*` / `*.v2.*` / 无后缀版 | 早期版本，保留备查 |
