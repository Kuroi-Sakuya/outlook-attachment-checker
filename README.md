# 附件检查助手 (Outlook Attachment Checker)

发送邮件时，如果正文提到了「附件 / 随函附送 / attached」等字样，但还没有添加附件，
就在发送前弹出提醒。所有检查都在本机 Outlook 内完成，邮件内容不会上传。

---

## 为什么 Outlook 会报“无法加载这个加载项”

这是**托管地址**的问题，不是代码的问题。

Outlook 加载项真正“被加载”的，不是这个仓库里的文件，而是**清单 (manifest.xml)**
里写的那些 `https://...` 网址。Outlook 会去这些网址下载 `commands.v3.html`、
`launchevent.v3.js` 来运行。

如果清单里的网址指向下面这两种地址，浏览器/Outlook 就**拒绝把它当网页或脚本运行**，
于是报“无法加载这个加载项”：

| 错误的托管地址 | 服务器返回的类型 | 结果 |
| --- | --- | --- |
| `raw.githubusercontent.com/.../launchevent.v3.js` | `Content-Type: text/plain` + `X-Content-Type-Options: nosniff` + `Content-Security-Policy: sandbox` | 浏览器禁止执行，加载失败 |
| `github.com/.../blob/.../commands.v3.html` | 返回的是 GitHub 的网页外壳，不是文件本身 | 内容不对，加载失败 |

> 实测本仓库的 raw 地址返回的就是 `content-type: text/plain` + `nosniff` + `sandbox`，
> 这正是“无法加载”的直接原因。

**正确做法：用一个会以正确 MIME 类型、通过 HTTPS 提供静态网页的主机。**
最简单免费的就是 **GitHub Pages**（地址形如 `https://<用户名>.github.io/<仓库名>/...`）。

---

## 修复步骤

### 1. 开启 GitHub Pages
1. 打开仓库 → **Settings** → 左侧 **Pages**
2. **Source** 选 **Deploy from a branch**
3. **Branch** 选 `main`，文件夹选 `/ (root)`，点 **Save**
4. 等 1–2 分钟，页面顶部会显示站点地址：
   `https://kuroi-sakuya.github.io/outlook-attachment-checker/`

### 2. 确认文件能直接打开
在浏览器里逐个打开下面三个地址，应当能看到网页/脚本内容（而不是下载、也不是 404）：

- https://kuroi-sakuya.github.io/outlook-attachment-checker/commands.v3.html
- https://kuroi-sakuya.github.io/outlook-attachment-checker/launchevent.v3.js
- https://kuroi-sakuya.github.io/outlook-attachment-checker/taskpane.html

### 3. 用本仓库的 `manifest.xml` 重新提交
- 本仓库已附带一份配置好 GitHub Pages 地址的 `manifest.xml`。
- 如果管理员之前已经提交过一版清单，请把 `manifest.xml` 里的 `<Id>` 换成**原来那一版的同一个 GUID**，
  这样在 Microsoft 365 管理中心里是“更新”而不是“新增”。
- 在 **Microsoft 365 管理中心 → 设置 → 集成应用 (Integrated apps)** 重新上传/更新这份清单。

### 4. 在 Outlook 里验证
- 关闭并重新打开 Outlook（集中部署的更新最多可能要等几十分钟到 24 小时才生效）。
- 新建一封邮件，正文写“见附件”但不加附件，点发送，应当弹出提醒。

---

## 技术要点

- 这是**发送时事件 (Smart Alerts / `OnMessageSend`)** 类型的加载项，要求邮箱满足
  **Mailbox 需求集 1.12**（较新的 Outlook 才支持）。
- 处理函数是 `launchevent.v3.js` 里的 `onMessageSendHandler`，通过
  `Office.actions.associate("onMessageSendHandler", ...)` 注册。
- `SendMode="SoftBlock"`：提醒后用户仍可选择继续发送。
- 托管主机的硬性要求：**HTTPS**、**正确的 MIME 类型**（.html → `text/html`，
  .js → `text/javascript`）、**不带 `nosniff`/`sandbox` 的封锁**。GitHub Pages 满足这些。

## 文件说明

| 文件 | 作用 |
| --- | --- |
| `manifest.xml` | 加载项清单（提交给 Outlook/管理中心的就是它） |
| `commands.v3.html` | 事件运行时承载页，加载 office.js 与处理脚本 |
| `launchevent.v3.js` | 发送时检查逻辑（加固 + 日志版，当前使用） |
| `taskpane.html` | 信息展示页 |
| `assets/` | 图标 |
| `*.v2.*` / 无后缀版 | 早期版本，保留备查 |
