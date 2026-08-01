(function () {
  "use strict";

  var cfg = window.FOLLOWUP_CONFIG || {};

  Office.onReady(function (info) {
    if (info.host !== Office.HostType.Outlook) return;
    try { populatePreview(); } catch (e) { /* preview is best-effort */ }
    document.getElementById("followupForm").addEventListener("submit", onSubmit);
  });

  // Read the selected email and show a short preview.
  function readItem() {
    var item = Office.context.mailbox.item;
    var restId = "";
    try {
      if (item.itemId) {
        restId = Office.context.mailbox.convertToRestId(
          item.itemId, Office.MailboxEnums.RestVersion.v2_0);
      }
    } catch (e) { restId = ""; }

    var to = (item.to || []).map(function (r) { return r.emailAddress; }).filter(Boolean);
    var sender = item.from ? item.from.emailAddress : (item.sender ? item.sender.emailAddress : "");
    var sent = "";
    try { if (item.dateTimeCreated) sent = new Date(item.dateTimeCreated).toISOString(); } catch (e) {}

    return {
      subject: item.subject || "",
      conversationId: item.conversationId || "",
      internetMessageId: item.internetMessageId || "",
      messageId: restId,
      sender: sender || "",
      recipients: to.join("; "),
      sentDateTime: sent
    };
  }

  function populatePreview() {
    var d = readItem();
    setText("pv-subject", d.subject || "—");
    setText("pv-to", d.recipients || "—");
    setText("pv-sent", d.sentDateTime ? d.sentDateTime.replace("T", " ").replace(/\..*/, "") : "—");
  }

  function onSubmit(ev) {
    ev.preventDefault();
    var btn = document.getElementById("submitBtn");

    var projectNumber = val("projectNumber").trim();
    if (!projectNumber) { showStatus("err", "נא להזין מספר פרויקט."); return; }

    var data = readItem();
    if (!data.internetMessageId) {
      showStatus("err", "לא ניתן לקרוא את מזהה המייל. פִּתחו את ההודעה ונסו שוב.");
      return;
    }

    var payload = {
      subject: data.subject,
      internetMessageId: data.internetMessageId,
      conversationId: data.conversationId,
      messageId: data.messageId,
      sender: data.sender,
      recipients: data.recipients,
      sentDateTime: data.sentDateTime,
      projectNumber: projectNumber,
      reminderDays: intOr("reminderDays", cfg.DEFAULT_REMINDER_DAYS || 3),
      maxReminders: intOr("maxReminders", cfg.DEFAULT_MAX_REMINDERS || 3),
      priority: parseInt(val("priority"), 10) || (cfg.DEFAULT_PRIORITY || 126760001)
    };

    btn.disabled = true;
    showStatus("info", "יוצר מעקב…");

    fetch(cfg.FLOW_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    }).then(function (res) {
      return res.text().then(function (t) {
        var body = {};
        try { body = t ? JSON.parse(t) : {}; } catch (e) { body = { raw: t }; }
        return { ok: res.ok, status: res.status, body: body };
      });
    }).then(function (r) {
      btn.disabled = false;
      if (r.ok && r.body && r.body.status === "ok") {
        var due = r.body.dueDate ? String(r.body.dueDate).substring(0, 10) : "";
        showStatus("ok", "✔ נוצר מעקב עבור: " + (r.body.name || projectNumber) +
          (due ? " · תזכורת ראשונה עד " + due : ""));
        document.getElementById("followupForm").reset();
        document.getElementById("reminderDays").value = cfg.DEFAULT_REMINDER_DAYS || 3;
        document.getElementById("maxReminders").value = cfg.DEFAULT_MAX_REMINDERS || 3;
      } else {
        showStatus("err", codeToMessage(r.body && r.body.code, r.status, r.body && r.body.message));
      }
    }).catch(function (err) {
      btn.disabled = false;
      showStatus("err", "שגיאת רשת/הרשאה בקריאה ל-Flow. בדקו חיבור/CORS. " + (err && err.message ? err.message : ""));
    });
  }

  function codeToMessage(code, httpStatus, serverMsg) {
    switch (code) {
      case "PROJECT_NOT_FOUND": return "מספר הפרויקט לא נמצא ב-PP_Projects. לא נוצר מעקב.";
      case "DUPLICATE":         return "כבר קיים מעקב עבור מייל זה.";
      case "MESSAGE_NOT_FOUND": return "המייל לא אותר בתיבת הדואר.";
      default: return serverMsg || ("יצירת המעקב נכשלה (HTTP " + httpStatus + ").");
    }
  }

  // helpers
  function val(id) { var el = document.getElementById(id); return el ? el.value : ""; }
  function intOr(id, def) { var n = parseInt(val(id), 10); return (isNaN(n) || n < 1) ? def : n; }
  function setText(id, t) { var el = document.getElementById(id); if (el) el.textContent = t; }
  function showStatus(kind, msg) {
    var el = document.getElementById("status");
    el.className = "status show " + kind;
    el.textContent = msg;
  }
})();
