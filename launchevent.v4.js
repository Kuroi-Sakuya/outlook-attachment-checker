/*
 * 附件检查助手 (Attachment Checker) —— 发送时事件处理 v4（加固版）
 *
 * 针对线上排查到的真实症状做了三处关键加固：
 *   1. 多函数名注册：无论部署到管理中心的清单里 FunctionName 写成什么，
 *      都能命中本处理函数（解决“注册了却没被调用”的对不上问题）。
 *   2. 硬性兜底超时：最多 HARD_TIMEOUT_MS 毫秒内一定调用 event.completed，
 *      绝不再卡死在“正在处理邮件…超过预期”。
 *   3. 单独给 body.getAsync / getAttachmentsAsync 各加一个短超时，
 *      任何一步挂住都不会拖垮整个发送流程。
 *
 * 日志横幅改为 “>>> ACHK v4:”，方便确认新文件是否真正生效。
 */

var VER = "v4";
var HARD_TIMEOUT_MS = 3500; // 整体兜底：超过即放行
var STEP_TIMEOUT_MS = 2500; // 单个异步 Office 调用的最长等待

function LOG(msg, extra) {
  try {
    if (extra !== undefined) {
      console.log(">>> ACHK " + VER + ": " + msg, extra);
    } else {
      console.log(">>> ACHK " + VER + ": " + msg);
    }
  } catch (e) {}
}

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

  function settle(options) {
    if (settled) {
      return;
    }
    settled = true;
    LOG("即将调用 event.completed，allowEvent=", options.allowEvent);
    try {
      event.completed(options);
      LOG("event.completed 调用成功");
    } catch (e) {
      LOG("event.completed 抛错: " + (e && e.message));
    }
  }

  // 1) 整体硬性兜底：无论下面发生什么，到点就放行。
  try {
    setTimeout(function () {
      LOG("硬性兜底超时 → 放行");
      settle({ allowEvent: true });
    }, HARD_TIMEOUT_MS);
  } catch (e) {
    LOG("布置兜底超时失败: " + (e && e.message));
  }

  // 2) 读正文（带单步超时保护）
  try {
    LOG("开始读取正文 body.getAsync");
    var bodyDone = false;
    setTimeout(function () {
      if (!bodyDone) {
        LOG("body.getAsync 单步超时 → 放行");
        settle({ allowEvent: true });
      }
    }, STEP_TIMEOUT_MS);

    Office.context.mailbox.item.body.getAsync("text", function (bodyResult) {
      bodyDone = true;
      try {
        LOG("body.getAsync 回调，status=", bodyResult && bodyResult.status);
        if (!bodyResult ||
            bodyResult.status !== Office.AsyncResultStatus.Succeeded ||
            bodyResult.value == null) {
          settle({ allowEvent: true });
          return;
        }

        if (!hasAttachmentMention(bodyResult.value)) {
          LOG("未命中附件关键词 → 放行");
          settle({ allowEvent: true });
          return;
        }

        LOG("命中关键词，开始读取附件 getAttachmentsAsync");
        var attDone = false;
        setTimeout(function () {
          if (!attDone) {
            LOG("getAttachmentsAsync 单步超时 → 放行");
            settle({ allowEvent: true });
          }
        }, STEP_TIMEOUT_MS);

        Office.context.mailbox.item.getAttachmentsAsync(function (attResult) {
          attDone = true;
          try {
            LOG("getAttachmentsAsync 回调，status=", attResult && attResult.status);
            if (!attResult || attResult.status !== Office.AsyncResultStatus.Succeeded) {
              settle({ allowEvent: true });
              return;
            }
            var atts = attResult.value || [];
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

// 注册处理函数。为防止部署的清单里 FunctionName 与代码不一致，
// 用多个常见候选名都关联到同一处理函数（多注册无害）。
var HANDLER_NAMES = [
  "onMessageSendHandler", // 与本仓库 manifest.xml 一致（首选）
  "onMessageSend",
  "checkAttachments",
  "validateBody",
  "attachmentCheck"
];
(function associateAll() {
  var ok = [];
  for (var i = 0; i < HANDLER_NAMES.length; i++) {
    try {
      Office.actions.associate(HANDLER_NAMES[i], onMessageSendHandler);
      ok.push(HANDLER_NAMES[i]);
    } catch (e) {}
  }
  LOG("脚本已加载，已注册候选函数名: " + ok.join(", "));
})();
