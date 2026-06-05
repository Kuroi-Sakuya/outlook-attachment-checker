/*
 * 附件检查助手 (Attachment Checker) —— 发送时事件处理
 *
 * 作用：发邮件前，自动检查正文里是否出现“附件 / 随函附送 / attached”等字样；
 *       若提到了却没有真正的附件，则弹出提醒，让你确认是否补充附件。
 *
 * 隐私：所有判断都在你本机的 Outlook 内运行，邮件内容不会上传到任何地方。
 *
 * 兼容：使用 ES5 写法、无 import，确保各平台 Outlook 运行时均可加载。
 */

// 中文触发词（直接子串匹配）。包含“附件”二字即可覆盖“见附件/附件如下/查收附件”等绝大多数说法。
var ZH_TERMS = [
  "附件", "附上", "附呈", "附寄", "附图", "附表",
  "随函附", "随附", "随文附", "兹附", "检附",
  "请查收", "附在邮件", "附于此", "见附", "附件为"
];

// 英文触发词（按单词边界、不区分大小写匹配）。
var EN_TERMS = [
  "attached", "attachment", "attaching",
  "enclosed", "enclosure",
  "please find", "pfa", "herewith"
];

// 回复/转发分隔标记：扫描关键词时只看这些标记之前“你新写的内容”，
// 避免被下方引用的历史邮件里的“附件”字样误触发。
var REPLY_MARKERS = [
  "发件人:", "发件人：", "寄件者", "发送时间:", "发送时间：",
  "From:", "Sent:", "-----Original Message-----", "-----原始邮件-----",
  "原始邮件", "原邮件", "________________________________"
];

function onMessageSendHandler(event) {
  Office.context.mailbox.item.body.getAsync(
    "text",
    { asyncContext: event },
    getBodyCallback
  );
}

function getBodyCallback(asyncResult) {
  var event = asyncResult.asyncContext;
  var body = "";
  if (asyncResult.status !== Office.AsyncResultStatus.Failed && asyncResult.value !== undefined) {
    body = asyncResult.value;
  } else {
    // 读取正文失败：放行，绝不因插件自身问题挡住发信。
    event.completed({ allowEvent: true });
    return;
  }

  if (hasAttachmentMention(body)) {
    Office.context.mailbox.item.getAttachmentsAsync(
      { asyncContext: event },
      getAttachmentsCallback
    );
  } else {
    event.completed({ allowEvent: true });
  }
}

// 截取“你新写的内容”（去掉下方引用的历史邮件），降低回复/转发时的误报。
function getComposedPortion(body) {
  if (!body) {
    return "";
  }
  var cutAt = body.length;
  var i, idx;

  for (i = 0; i < REPLY_MARKERS.length; i++) {
    idx = body.indexOf(REPLY_MARKERS[i]);
    if (idx !== -1 && idx < cutAt) {
      cutAt = idx;
    }
  }

  // 中文“在 …… 写道：” / 英文“On …… wrote:”
  var zhWrote = body.match(/在[\s\S]{0,80}?写道[:：]/);
  if (zhWrote && zhWrote.index < cutAt) {
    cutAt = zhWrote.index;
  }
  var enWrote = body.match(/On[\s\S]{0,120}?wrote:/);
  if (enWrote && enWrote.index < cutAt) {
    cutAt = enWrote.index;
  }

  return body.substring(0, cutAt);
}

function hasAttachmentMention(fullBody) {
  var text = getComposedPortion(fullBody);
  if (!text) {
    return false;
  }

  // 中文：直接子串匹配
  var i;
  for (i = 0; i < ZH_TERMS.length; i++) {
    if (text.indexOf(ZH_TERMS[i]) !== -1) {
      return true;
    }
  }

  // 英文：按单词边界、忽略大小写
  var lower = text.toLowerCase();
  for (i = 0; i < EN_TERMS.length; i++) {
    var safe = EN_TERMS[i].replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    var re = new RegExp("\\b" + safe + "\\b", "i");
    if (re.test(lower)) {
      return true;
    }
  }

  return false;
}

function getAttachmentsCallback(asyncResult) {
  var event = asyncResult.asyncContext;

  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    // 读取附件列表失败：放行，避免误挡。
    event.completed({ allowEvent: true });
    return;
  }

  var attachments = asyncResult.value || [];

  // 只要存在一个“非内嵌”的真实附件，就放行。
  var i;
  for (i = 0; i < attachments.length; i++) {
    if (attachments[i].isInline === false) {
      event.completed({ allowEvent: true });
      return;
    }
  }

  // 提到了附件，但没有真实附件 → 提醒（PromptUser：可选择“仍要发送”）。
  event.completed({
    allowEvent: false,
    errorMessage: "邮件正文里提到了「附件 / 随函附送」等字样，但似乎还没有添加附件。要先添加附件再发送吗？"
  });
}

// 必须：把清单里声明的处理函数名（onMessageSendHandler）映射到此实现，否则发送会被卡住。
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
