/*
 * English Reply Headers for Outlook
 * Target: new Outlook on Windows / Outlook on the web
 *
 * HTML replies:
 *   Only the newest Outlook-generated reply header block (divRplyFwdMsg)
 *   is modified. Older quoted message headers are left untouched.
 *
 * Plain-text replies:
 *   Only a strict Outlook-style header sequence is modified:
 *   差出人 -> 送信 -> 宛先 -> [CC] -> 件名
 */

(function (root) {
  "use strict";

  var CONFIG = {
    addOriginalMessageSeparator: true,
    convertJapaneseSentDate: true,
    convertOnlyNewestHeader: true,
    retryDelaysMs: [0, 80, 160, 280, 450, 700, 1000]
  };

  var MONTHS = [
    "Jan", "Feb", "Mar", "Apr", "May", "Jun",
    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
  ];

  var LABELS = {
    from: "From:",
    sent: "Sent:",
    to: "To:",
    cc: "Cc:",
    subject: "Subject:"
  };

  function normalizeWhitespace(value) {
    return String(value == null ? "" : value)
      .replace(/\u00a0/g, " ")
      .replace(/[ \t]+/g, " ")
      .trim();
  }

  function normalizeLabel(value) {
    return normalizeWhitespace(value)
      .replace(/[：:]\s*$/, "")
      .replace(/\s+/g, "")
      .toLowerCase();
  }

  function labelKey(value) {
    var label = normalizeLabel(value);
    if (label === "差出人") return "from";
    if (label === "送信") return "sent";
    if (label === "宛先") return "to";
    if (label === "件名") return "subject";
    if (label === "from") return "from";
    if (label === "sent") return "sent";
    if (label === "to") return "to";
    if (label === "subject") return "subject";
    if (label === "cc" || label === "ｃｃ") return "cc";
    return null;
  }

  function convertJapaneseDate(value) {
    var original = normalizeWhitespace(value);
    if (!CONFIG.convertJapaneseSentDate) return original;

    // Typical new Outlook Japanese form:
    // 2026 年 9 月 11 日 (金曜日) 14:05
    // Also accepts full-width parentheses and optional seconds.
    var m = original.match(
      /^(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日(?:\s*[（(][^）)]*[）)])?\s*(\d{1,2}):(\d{2})(?::(\d{2}))?$/
    );
    if (!m) return original;

    var year = Number(m[1]);
    var month = Number(m[2]);
    var day = Number(m[3]);
    var hour = String(Number(m[4])).padStart(2, "0");
    var minute = m[5];
    var second = m[6] ? ":" + m[6] : "";

    if (month < 1 || month > 12 || day < 1 || day > 31 || Number(m[4]) > 23 || Number(minute) > 59 || Number(m[6] || 0) > 59) return original;
    if (new Date(Date.UTC(year, month - 1, day)).getUTCDate() !== day) return original;
    return day + " " + MONTHS[month - 1] + " " + year + " " + hour + ":" + minute + second;
  }

  function delay(ms) {
    return new Promise(function (resolve) { setTimeout(resolve, ms); });
  }

  function isWhitespaceTextNode(node) {
    return node && node.nodeType === 3 && !normalizeWhitespace(node.nodeValue || "");
  }

  function previousMeaningfulSibling(node) {
    var prev = node ? node.previousSibling : null;
    while (prev && isWhitespaceTextNode(prev)) prev = prev.previousSibling;
    return prev;
  }

  function hasSeparatorImmediatelyBefore(block) {
    var prev = previousMeaningfulSibling(block);
    if (!prev) return false;
    if (prev.nodeType === 1 && prev.getAttribute && prev.getAttribute("data-erh-separator") === "1") return true;
    return /-{3,}\s*Original Message\s*-{3,}/i.test(prev.textContent || "");
  }

  function addHtmlSeparator(block) {
    if (!CONFIG.addOriginalMessageSeparator || !block || !block.parentNode) return false;
    if (hasSeparatorImmediatelyBefore(block)) return false;

    var doc = block.ownerDocument;
    var sep = doc.createElement("div");
    sep.setAttribute("data-erh-separator", "1");
    sep.textContent = "----- Original Message -----";
    block.parentNode.insertBefore(sep, block);
    return true;
  }

  function nodesUntilBreak(labelElement) {
    var nodes = [];
    var node = labelElement ? labelElement.nextSibling : null;
    var guard = 0;
    while (node && guard < 32) {
      if (node.nodeType === 1 && String(node.tagName).toUpperCase() === "BR") break;
      nodes.push(node);
      node = node.nextSibling;
      guard += 1;
    }
    return nodes;
  }

  function nodeText(node) {
    if (!node) return "";
    if (node.nodeType === 3) return node.nodeValue || "";
    return node.textContent || "";
  }

  function rewriteSentDateAfterLabel(labelElement) {
    if (!CONFIG.convertJapaneseSentDate) return false;
    var nodes = nodesUntilBreak(labelElement);
    if (!nodes.length) return false;

    var original = "";
    for (var i = 0; i < nodes.length; i += 1) original += nodeText(nodes[i]);
    var converted = convertJapaneseDate(original);
    if (!converted || converted === normalizeWhitespace(original)) return false;

    var doc = labelElement.ownerDocument;
    var replacement = doc.createTextNode(" " + converted);
    var parent = labelElement.parentNode;
    parent.insertBefore(replacement, nodes[0]);
    for (var j = 0; j < nodes.length; j += 1) {
      if (nodes[j].parentNode === parent) parent.removeChild(nodes[j]);
    }
    return true;
  }

  function isReplyHeaderId(value) {
    var id = String(value || "");
    return /^(?:x_)?divRplyFwdMsg(?:_\d+)?$/i.test(id);
  }

  function findHtmlReplyHeaderBlocks(doc) {
    if (!doc || !doc.querySelectorAll) return [];
    var all = Array.prototype.slice.call(doc.querySelectorAll("[id]"));
    // DOM order is authoritative, even when a newer header has a prefixed ID.
    var blocks = all.filter(function (el) { return isReplyHeaderId(el.id); });
    if (CONFIG.convertOnlyNewestHeader && blocks.length > 1) return [blocks[0]];
    return blocks;
  }

  function translateHtmlBlock(block) {
    if (!block || !block.querySelectorAll) return false;
    var changed = false;
    var candidates = Array.prototype.slice.call(block.querySelectorAll("b, strong"));
    candidates = candidates.filter(function (el) {
      var parent = el.parentElement;
      while (parent && parent !== block) {
        if (isReplyHeaderId(parent.id)) return false;
        parent = parent.parentElement;
      }
      return !!labelKey(el.textContent);
    });
    var keys = candidates.map(function (el) { return labelKey(el.textContent); }).join(",");
    if (keys !== "from,sent,to,subject" && keys !== "from,sent,to,cc,subject") return false;
    // An English newest block stops processing; never continue into older mail.
    if (!candidates.some(function (el) { return /差出人|送信|宛先|件名/.test(el.textContent); })) return false;

    for (var i = 0; i < candidates.length; i += 1) {
      var el = candidates[i];
      var key = labelKey(el.textContent || "");
      if (!key) continue;

      if (key === "sent" && rewriteSentDateAfterLabel(el)) changed = true;

      if (normalizeWhitespace(el.textContent) !== LABELS[key]) {
        el.textContent = LABELS[key];
        changed = true;
      }
    }

    if (addHtmlSeparator(block)) changed = true;
    return changed;
  }

  function replaceHtmlHeaders(input) {
    var html = String(input == null ? "" : input);

    if (typeof DOMParser === "undefined") {
      // This path is mainly for non-browser test/runtime environments.
      // Be conservative: don't globally replace labels in arbitrary HTML.
      return { text: html, changed: false, reason: "DOMParser unavailable." };
    }

    var parser = new DOMParser();
    var doc = parser.parseFromString(html, "text/html");
    var blocks = findHtmlReplyHeaderBlocks(doc);
    if (!blocks.length) return { text: html, changed: false, reason: "No divRplyFwdMsg block found." };

    var changed = false;
    for (var i = 0; i < blocks.length; i += 1) {
      if (translateHtmlBlock(blocks[i])) changed = true;
    }

    return {
      text: changed ? doc.body.innerHTML : html,
      changed: changed,
      reason: changed ? "Converted Outlook reply header block." :
        (/差出人|送信|宛先|件名/.test(blocks[0].textContent) ? "Unsupported Outlook header markup." : "Reply header block was already converted.")
    };
  }

  function parsePlainHeaderLine(line) {
    var m = String(line).match(/^([ \t]*)(差出人|送信|宛先|件名|From|Sent|To|Subject|CC|ＣＣ|ｃｃ)\s*[:：][ \t]*(.*)$/i);
    if (!m) return null;
    var key = labelKey(m[2]);
    if (!key) return null;
    return { indent: m[1], key: key, value: m[3] };
  }

  function findPlainHeaderBlock(lines) {
    for (var start = 0; start < lines.length; start += 1) {
      var first = parsePlainHeaderLine(lines[start]);
      if (!first || first.key !== "from") continue;

      var found = { from: start };
      var last = start;
      var limit = Math.min(lines.length, start + 14);

      for (var i = start + 1; i < limit; i += 1) {
        var parsed = parsePlainHeaderLine(lines[i]);
        if (!parsed) {
          if (!lines[i].trim()) continue;
          break;
        }

        if (parsed.key === "sent" && found.sent == null) {
          found.sent = i;
          last = i;
        } else if (parsed.key === "to" && found.sent != null && found.to == null) {
          found.to = i;
          last = i;
        } else if (parsed.key === "cc" && found.to != null && found.subject == null) {
          found.cc = i;
          last = i;
        } else if (parsed.key === "subject" && found.to != null) {
          found.subject = i;
          last = i;
          break;
        } else if (parsed.key === "from" && i !== start) {
          break;
        } else {
          break;
        }
      }

      if (found.sent != null && found.to != null && found.subject != null) {
        found.end = last;
        return found;
      }
    }
    return null;
  }

  function plainSeparatorAlreadyBefore(lines, index) {
    var from = Math.max(0, index - 3);
    for (var i = from; i < index; i += 1) {
      if (/-{3,}\s*Original Message\s*-{3,}/i.test(lines[i])) return true;
    }
    return false;
  }

  function rewritePlainLine(line) {
    var parsed = parsePlainHeaderLine(line);
    if (!parsed) return line;
    var value = parsed.key === "sent" ? convertJapaneseDate(parsed.value) : parsed.value;
    return parsed.indent + LABELS[parsed.key] + (value ? " " + value : "");
  }

  function replacePlainTextHeaders(input) {
    var text = String(input == null ? "" : input);
    var eol = text.indexOf("\r\n") !== -1 ? "\r\n" : "\n";
    var lines = text.split(/\r?\n/);
    var block = findPlainHeaderBlock(lines);
    if (!block) return { text: text, changed: false, reason: "No strict Outlook reply header sequence found." };
    if (![block.from, block.sent, block.to, block.subject].some(function (i) { return /^\s*(差出人|送信|宛先|件名)\s*[:：]/.test(lines[i]); })) {
      return { text: text, changed: false, reason: "Reply header was already converted." };
    }

    var indexes = [block.from, block.sent, block.to, block.cc, block.subject];
    var changed = false;
    for (var i = 0; i < indexes.length; i += 1) {
      var idx = indexes[i];
      if (idx == null) continue;
      var next = rewritePlainLine(lines[idx]);
      if (next !== lines[idx]) {
        lines[idx] = next;
        changed = true;
      }
    }

    if (CONFIG.addOriginalMessageSeparator && !plainSeparatorAlreadyBefore(lines, block.from)) {
      var indent = parsePlainHeaderLine(lines[block.from]) ? parsePlainHeaderLine(lines[block.from]).indent : "";
      lines.splice(block.from, 0, indent + "----- Original Message -----");
      changed = true;
    }

    return {
      text: changed ? lines.join(eol) : text,
      changed: changed,
      reason: changed ? "Converted strict plain-text reply header block." : "Reply header was already converted."
    };
  }

  function getComposeInfo(item) {
    return new Promise(function (resolve) {
      if (!item || typeof item.getComposeTypeAsync !== "function") {
        resolve(null);
        return;
      }
      item.getComposeTypeAsync(function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value || null);
        else resolve(null);
      });
    });
  }

  function getBodyType(item) {
    return new Promise(function (resolve, reject) {
      item.body.getTypeAsync(function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
        else reject(result.error);
      });
    });
  }

  function getBody(item, bodyType, mode) {
    return new Promise(function (resolve, reject) {
      item.body.getAsync(bodyType, { bodyMode: mode || Office.MailboxEnums.BodyMode.HostConfig }, function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
        else reject(result.error);
      });
    });
  }

  function setBody(item, bodyType, value) {
    return new Promise(function (resolve, reject) {
      item.body.setAsync(
        value,
        { coercionType: bodyType, bodyMode: Office.MailboxEnums.BodyMode.HostConfig },
        function (result) {
          if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
          else reject(result.error);
        }
      );
    });
  }

  function isReplyOrForward(composeInfo) {
    if (!composeInfo || !composeInfo.composeType) return true;
    return composeInfo.composeType === "reply" || composeInfo.composeType === "forward";
  }

  async function transformCurrentItem() {
    var item = Office.context && Office.context.mailbox && Office.context.mailbox.item;
    if (!item || !item.body) return { changed: false, reason: "No compose item." };

    var composeInfo = await getComposeInfo(item);
    if (!isReplyOrForward(composeInfo)) {
      return { changed: false, reason: "New message; nothing to convert." };
    }

    var bodyType = composeInfo && composeInfo.coercionType ? composeInfo.coercionType : await getBodyType(item);
    var body = await getBody(item, bodyType);
    var isHtml = bodyType === Office.CoercionType.Html || String(bodyType).toLowerCase() === "html";
    var transformed = isHtml ? replaceHtmlHeaders(body) : replacePlainTextHeaders(body);

    if (!transformed.changed) return transformed;
    // Avoid overwriting edits that arrived while the first read was in flight.
    if (await getBody(item, bodyType) !== body) return { changed: false, reason: "Body changed during conversion; try again." };
    await setBody(item, bodyType, transformed.text);
    return transformed;
  }

  async function transformWithRetry() {
    var last = { changed: false, reason: "No reply header found." };
    for (var i = 0; i < CONFIG.retryDelaysMs.length; i += 1) {
      var wait = CONFIG.retryDelaysMs[i] - (i ? CONFIG.retryDelaysMs[i - 1] : 0);
      if (wait) await delay(wait);
      try {
        last = await transformCurrentItem();
        if (last.changed) return last;
        if (/already converted|Body changed/.test(last.reason)) return last;
        if (last.reason === "New message; nothing to convert.") return last;
      } catch (err) {
        last = { changed: false, reason: "Office API error.", errorCode: err && (err.code || err.name) || "Unknown" };
        return last;
      }
    }
    return last;
  }

  async function onNewMessageComposeHandler(event) {
    try {
      await transformWithRetry();
    } catch (err) {
      if (typeof console !== "undefined" && console.error) console.error("English Reply Headers:", err);
    } finally {
      if (event && event.completed) event.completed();
    }
  }

  // Status only: never include body text, addresses or subjects in diagnostics.
  function showStatus(message, failed) {
    return new Promise(function (resolve) {
      var done = false;
      var timer = setTimeout(finish, 2500);
      function finish() { if (!done) { done = true; clearTimeout(timer); resolve(); } }
      try {
        var item = Office.context.mailbox.item;
        if (!item.notificationMessages) { finish(); return; }
        var details = { type: failed ? "errorMessage" : "informationalMessage", message: ("ERH 1.3.1: " + message).slice(0, 150) };
        if (!failed) { details.icon = "Icon.16"; details.persistent = true; }
        item.notificationMessages.replaceAsync("erh-status", details, finish);
      } catch (_) { finish(); }
    });
  }

  async function describeResult(result) {
    if (result.changed) return "英語ヘッダーへ変換しました。";
    if (/already converted/.test(result.reason)) return "取得した最新ヘッダーは英語です。";
    if (/New message/.test(result.reason)) return "新規メールと判定されました。返信・転送画面で実行してください。";
    if (/Body changed/.test(result.reason)) return "処理中に本文が変わったため中止しました。もう一度押してください。";
    if (result.errorCode) return "APIエラー: " + String(result.errorCode).replace(/[^A-Za-z0-9_-]/g, "").slice(0, 45);
    if (/DOMParser/.test(result.reason)) return "HTML解析機能が利用できません。[NO_DOM]";
    if (/Unsupported/.test(result.reason)) return "ヘッダーはありますが、想定と異なるHTML形式です。[MARKUP]";
    var item = Office.context.mailbox.item;
    if (!item || !item.body) return "メールの作成画面が取得できません。[NO_ITEM]";
    var type = await getBodyType(item);
    var current = await getBody(item, type);
    var full = await getBody(item, type, Office.MailboxEnums.BodyMode.FullBody);
    // FullBody is read only. Do not replace the conversation as a fallback.
    var probe = String(type).toLowerCase() === "html" ? replaceHtmlHeaders(full) : replacePlainTextHeaders(full);
    if (full !== current && probe.changed) return "引用ヘッダーが取得範囲の外にあります。[SCOPE]";
    return "対応する日本語ヘッダーが見つかりません。[NO_HEADER " + String(type) + ", " + current.length + "文字]";
  }

  async function manualConvert(event) {
    try {
      await showStatus("起動しました。ヘッダーを確認中…", false);
      var result = await transformWithRetry();
      await showStatus(await describeResult(result), !!result.errorCode);
    } catch (err) {
      await showStatus("診断エラー: " + String(err && (err.code || err.name) || "Unknown").replace(/[^A-Za-z0-9_-]/g, "").slice(0, 45), true);
    } finally {
      if (event && event.completed) event.completed();
    }
  }

  root.EnglishReplyHeaders = {
    config: CONFIG,
    convertJapaneseDate: convertJapaneseDate,
    labelKey: labelKey,
    findPlainHeaderBlock: findPlainHeaderBlock,
    replacePlainTextHeaders: replacePlainTextHeaders,
    replaceHtmlHeaders: replaceHtmlHeaders,
    findHtmlReplyHeaderBlocks: findHtmlReplyHeaderBlocks,
    transformCurrentItem: transformCurrentItem,
    transformWithRetry: transformWithRetry,
    manualConvert: manualConvert
  };

  if (typeof Office !== "undefined" && Office.actions && Office.actions.associate) {
    Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
    Office.actions.associate("manualConvert", manualConvert);
  }
  if (typeof Office !== "undefined") Office.onReady(function () {});
})(typeof globalThis !== "undefined" ? globalThis : this);
