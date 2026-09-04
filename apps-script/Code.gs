/**
 * Build with Sri — assessment backend (Google Apps Script)
 *
 * Actions (POSTed as JSON): saveLead, sendOtp, verifyOtp
 * OTP is handled by Twilio Verify (WhatsApp channel, SMS fallback).
 *
 * REQUIRED Script Properties (Project Settings -> Script Properties):
 *   TWILIO_ACCOUNT_SID        - from Twilio Console home
 *   TWILIO_AUTH_TOKEN         - from Twilio Console home (keep secret)
 *   TWILIO_VERIFY_SERVICE_SID - the "VA..." SID of your Verify Service
 *   SHEET_ID                  - ID of the Google Sheet that stores leads
 *                               (the long string in the sheet's URL)
 */

var PROPS = PropertiesService.getScriptProperties();

var LEAD_HEADERS = [
  "leadId", "timestamp", "verified",
  "q1", "q2", "q3", "q4", "q5", "q6", "q7", "q8", "q9",
  "name", "email", "phone"
];

function doPost(e) {
  var body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOut({ ok: false, message: "Bad request." });
  }
  try {
    switch (body.action) {
      case "saveLead":  return saveLead(body.lead);
      case "sendOtp":   return sendOtp(body.leadId, body.phone);
      case "verifyOtp": return verifyOtp(body.leadId, body.otp);
      default:          return jsonOut({ ok: false, message: "Unknown action." });
    }
  } catch (err) {
    return jsonOut({ ok: false, message: err.message || "Server error." });
  }
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function leadsSheet() {
  var ss = SpreadsheetApp.openById(PROPS.getProperty("SHEET_ID"));
  var sh = ss.getSheetByName("Leads");
  if (!sh) {
    sh = ss.insertSheet("Leads");
    sh.appendRow(LEAD_HEADERS);
  }
  return sh;
}

/* ---------- saveLead ---------- */

function saveLead(lead) {
  if (!lead || !lead.email || !lead.phone) {
    return jsonOut({ ok: false, message: "Missing contact details." });
  }
  var leadId = Utilities.getUuid();
  leadsSheet().appendRow([
    leadId, new Date().toISOString(), "NO",
    lead.q1 || "", lead.q2 || "", lead.q3 || "", lead.q4 || "", lead.q5 || "",
    lead.q6 || "", lead.q7 || "", lead.q8 || "", lead.q9 || "",
    lead.name || "", lead.email || "", lead.phone || ""
  ]);
  return jsonOut({ ok: true, leadId: leadId });
}

/* ---------- OTP via Twilio Verify ---------- */

function normalizePhone(raw) {
  var s = String(raw || "").trim();
  var hasPlus = s.charAt(0) === "+";
  var digits = s.replace(/\D/g, "");
  if (!digits) throw new Error("Please enter a valid WhatsApp number.");
  if (hasPlus) return "+" + digits;
  digits = digits.replace(/^0+/, "");
  // A bare 10-digit number is assumed to be Indian (the form hints +91).
  if (digits.length === 10) return "+91" + digits;
  return "+" + digits;
}

function sendOtp(leadId, phone) {
  if (!leadId) return jsonOut({ ok: false, message: "Missing lead reference." });
  var to = normalizePhone(phone);

  // Remember the exact number we verified against for this lead.
  CacheService.getScriptCache().put("otp_phone_" + leadId, to, 1800);

  try {
    twilioVerify("Verifications", { To: to, Channel: "whatsapp" });
  } catch (err) {
    // WhatsApp can fail for numbers without WhatsApp — fall back to SMS.
    twilioVerify("Verifications", { To: to, Channel: "sms" });
  }
  return jsonOut({ ok: true });
}

function verifyOtp(leadId, otp) {
  if (!leadId) return jsonOut({ ok: false, message: "Missing lead reference." });
  var to = CacheService.getScriptCache().get("otp_phone_" + leadId);
  if (!to) {
    // Cache expired (30 min) — recover the phone from the sheet.
    to = normalizePhone(findLeadPhone(leadId));
  }

  var result = twilioVerify("VerificationCheck", { To: to, Code: String(otp || "") });
  if (result.status === "approved") {
    markVerified(leadId);
    return jsonOut({ ok: true, verified: true });
  }
  return jsonOut({ ok: true, verified: false, message: "That code is not valid. Please check and try again." });
}

function twilioVerify(path, params) {
  var sid = PROPS.getProperty("TWILIO_ACCOUNT_SID");
  var token = PROPS.getProperty("TWILIO_AUTH_TOKEN");
  var service = PROPS.getProperty("TWILIO_VERIFY_SERVICE_SID");
  if (!sid || !token || !service) {
    throw new Error("Verification service is not configured yet.");
  }
  var res = UrlFetchApp.fetch(
    "https://verify.twilio.com/v2/Services/" + service + "/" + path,
    {
      method: "post",
      payload: params,
      headers: { Authorization: "Basic " + Utilities.base64Encode(sid + ":" + token) },
      muteHttpExceptions: true
    }
  );
  var data = JSON.parse(res.getContentText());
  if (res.getResponseCode() >= 300) {
    // Twilio's message is developer-facing; keep the user-facing one generic.
    console.error("Twilio error: " + res.getContentText());
    throw new Error("We could not send the verification code. Please check the number and try again.");
  }
  return data;
}

/* ---------- sheet helpers ---------- */

function findLeadPhone(leadId) {
  var sh = leadsSheet();
  var values = sh.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (values[i][0] === leadId) return values[i][LEAD_HEADERS.indexOf("phone")];
  }
  throw new Error("We could not find your submission. Please start again.");
}

function markVerified(leadId) {
  var sh = leadsSheet();
  var values = sh.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (values[i][0] === leadId) {
      sh.getRange(i + 1, LEAD_HEADERS.indexOf("verified") + 1).setValue("YES");
      return;
    }
  }
}
