/*
 * 附件检查助手 (Attachment Checker) —— 发送时事件处理 v3（加固 + 日志版）
 *
 * 相比 v2 增加了详细 console.log 追踪，便于在浏览器控制台看到每一步，
 * 精确定位“一直处理中”卡在哪个环节。
 * 看门狗（最多 6 秒必放行）与全程 try/catch 兜底保留。
 */

var LOG = function (msg, extra) {
  try {
    if (extra !== undefined) {
      console.log(">>> ACHK v3: " + msg, extra);
    } else {
      console.log(">>> ACHK v3: " + msg);
    }
  } catch (e) {}
};

var ZH_TERMS = [
  "附件", "附上", "附呈", "附寄", "附图", "附表",
  "随函附", "随附", "随文附", "兹附", "检附",
  "请查收", "附在邮件", "附于此", "见附", "附件为"
];

var EN_TERMS = [
  "attached", "attachment", "attaching",
  "enclosed", "enclosure",
  "please find", "pfa", "herewith"
];

var REPLY_MARKERS = [
  "发件人:", "发件人：", "寄件者", "发送时间:", "发送时间：",
  "From:", "Sent:", "-----Original Message-----", "-----原始邮件-----",
  "原始邮件", "原邮件", "________________________________"
];

function onMessageSendHandler(event) {
  LOG("onMessageSendHandler 已触发");
  var settled = false;
  var watchdog = null;

  function settle(options) {
    if (settled) {
      LOG("settle 被重复调用，已忽略");
      return;
    }
    settled = true;
    if (watchdog) {
      try { clearTimeout(watchdog); } catch (e) {}
    }
    LOG("即将调用 event.completed，allowEvent=", options.allowEvent);
    try {
      event.completed(options);
      LOG("event.completed 调用成功");
    } catch (e) {
      LOG("event.completed 抛错: " + (e && e.message));
    }
  }

  try {
    watchdog = setTimeout(function () {
      LOG("看门狗超时(6s) → 放行");
      settle({ allowEvent: true });
    }, 6000);
    LOG("看门狗已布置(6s)");
  } catch (e) {
    LOG("布置看门狗失败: " + (e && e.message));
  }

  try {
    LOG("开始读取正文 body.getAsync");
    Office.context.mailbox.item.body.getAsync("text", {}, function (bodyResult) {
      try {
        LOG("body.getAsync 回调，status=", bodyResult && bodyResult.status);
        if (!bodyResult || bodyResult.status === Office.AsyncResultStatus.Failed || bodyResult.value == null) {
          LOG("读取正文失败/为空 → 放行");
          settle({ allowEvent: true });
          return;
        }

        LOG("正文长度=", String(bodyResult.value.length));
        var matched = hasAttachmentMention(bodyResult.value);
        LOG("是否命中附件关键词=", matched);
        if (!matched) {
          settle({ allowEvent: true });
          return;
        }

        LOG("开始读取附件 getAttachmentsAsync");
        Office.context.mailbox.item.getAttachmentsAsync({}, function (attResult) {
          try {
            LOG("getAttachmentsAsync 回调，status=", attResult && attResult.status);
            if (!attResult || attResult.status === Office.AsyncResultStatus.Failed) {
              LOG("读取附件失败 → 放行");
              settle({ allowEvent: true });
              return;
            }
            var atts = attResult.value || [];
            LOG("附件数量=", String(atts.length));
            for (var i = 0; i < atts.length; i++) {
              if (atts[i].isInline === false) {
                LOG("发现非内嵌附件 → 放行");
                settle({ allowEvent: true });
                return;
              }
            }
            LOG("命中关键词但无真实附件 → 拦截并提醒");
            settle({
              allowEvent: false,
              errorMessage: "邮件正文里提到了「附件 / 随函附送」等字样，但似乎还没有添加附件。要先添加附件再发送吗？"
            });
          } catch (e) {
            LOG("附件回调异常 → 放行: " + (e && e.message));
            settle({ allowEvent: true });
          }
        });
      } catch (e) {
        LOG("正文回调异常 → 放行: " + (e && e.message));
        settle({ allowEvent: true });
      }
    });
  } catch (e) {
    LOG("getAsync 调用异常 → 放行: " + (e && e.message));
    settle({ allowEvent: true });
  }
}

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
  try {
    var zhWrote = body.match(/在[\s\S]{0,80}?写道[:：]/);
    if (zhWrote && zhWrote.index < cutAt) {
      cutAt = zhWrote.index;
    }
    var enWrote = body.match(/On[\s\S]{0,120}?wrote:/);
    if (enWrote && enWrote.index < cutAt) {
      cutAt = enWrote.index;
    }
  } catch (e) {}
  return body.substring(0, cutAt);
}

function hasAttachmentMention(fullBody) {
  var text = getComposedPortion(fullBody);
  if (!text) {
    return false;
  }
  var i;
  for (i = 0; i < ZH_TERMS.length; i++) {
    if (text.indexOf(ZH_TERMS[i]) !== -1) {
      return true;
    }
  }
  var lower = text.toLowerCase();
  for (i = 0; i < EN_TERMS.length; i++) {
    try {
      var safe = EN_TERMS[i].replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      var re = new RegExp("\\b" + safe + "\\b", "i");
      if (re.test(lower)) {
        return true;
      }
    } catch (e) {}
  }
  return false;
}

// 映射处理函数；并在脚本加载时打印一条，便于确认“新代码确实加载了”。
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
LOG("脚本已加载，onMessageSendHandler 已关联完成");
