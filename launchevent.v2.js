/*
 * 附件检查助手 (Attachment Checker) —— 发送时事件处理 v2（加固版）
 *
 * 作用：发邮件前，自动检查正文里是否出现“附件 / 随函附送 / attached”等字样；
 *       若提到了却没有真正的附件，则弹出提醒。
 *
 * 加固要点（修复“一直转圈”问题）：
 *   1) 看门狗计时器：无论如何最多 6 秒内必定放行，绝不再卡住发送。
 *   2) 用闭包持有 event，不依赖 asyncContext 回传（新版 Mac Outlook 上可能丢失）。
 *   3) 全程 try/catch 兜底：任何异常都直接放行。
 *
 * 隐私：所有判断都在本机 Outlook 内运行，邮件内容不会上传到任何地方。
 */

// 中文触发词（直接子串匹配）。含“附件”即可覆盖“见附件/附件如下/查收附件”等绝大多数说法。
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

// 回复/转发分隔标记：扫描关键词时只看这些标记之前“你新写的内容”，避免被引用的历史邮件误触发。
var REPLY_MARKERS = [
  "发件人:", "发件人：", "寄件者", "发送时间:", "发送时间：",
  "From:", "Sent:", "-----Original Message-----", "-----原始邮件-----",
  "原始邮件", "原邮件", "________________________________"
];

function onMessageSendHandler(event) {
  var settled = false;
  var watchdog = null;

  // 统一出口：保证 event.completed 只被调用一次，并清除看门狗。
  function settle(options) {
    if (settled) {
      return;
    }
    settled = true;
    if (watchdog) {
      try { clearTimeout(watchdog); } catch (e) {}
    }
    try {
      event.completed(options);
    } catch (e) {
      // 忽略：极端情况下重复完成会报错，直接吞掉。
    }
  }

  // 看门狗：最多 6 秒，绝不让用户卡住（超时则放行发送）。
  try {
    watchdog = setTimeout(function () {
      settle({ allowEvent: true });
    }, 6000);
  } catch (e) {}

  try {
    Office.context.mailbox.item.body.getAsync("text", {}, function (bodyResult) {
      try {
        if (!bodyResult || bodyResult.status === Office.AsyncResultStatus.Failed || bodyResult.value == null) {
          settle({ allowEvent: true });
          return;
        }

        if (!hasAttachmentMention(bodyResult.value)) {
          settle({ allowEvent: true });
          return;
        }

        // 正文提到了附件 → 检查是否真有非内嵌附件。
        Office.context.mailbox.item.getAttachmentsAsync({}, function (attResult) {
          try {
            if (!attResult || attResult.status === Office.AsyncResultStatus.Failed) {
              settle({ allowEvent: true });
              return;
            }
            var atts = attResult.value || [];
            for (var i = 0; i < atts.length; i++) {
              if (atts[i].isInline === false) {
                settle({ allowEvent: true });
                return;
              }
            }
            // 提到了附件，但没有真实附件 → 提醒（PromptUser：可选择“仍要发送”）。
            settle({
              allowEvent: false,
              errorMessage: "邮件正文里提到了「附件 / 随函附送」等字样，但似乎还没有添加附件。要先添加附件再发送吗？"
            });
          } catch (e) {
            settle({ allowEvent: true });
          }
        });
      } catch (e) {
        settle({ allowEvent: true });
      }
    });
  } catch (e) {
    settle({ allowEvent: true });
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
  // 中文：直接子串匹配
  for (i = 0; i < ZH_TERMS.length; i++) {
    if (text.indexOf(ZH_TERMS[i]) !== -1) {
      return true;
    }
  }

  // 英文：按单词边界、忽略大小写
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

// 必须：把清单里声明的处理函数名映射到此实现。
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
