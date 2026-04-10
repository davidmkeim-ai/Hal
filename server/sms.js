import { config, isSmsConfigured } from "./config.js";

export function getSmsStatus() {
  return {
    smsEnabled: true,
    provider: config.sms.provider || "unconfigured",
    providerConfigured: isSmsConfigured(),
    fromNumber: config.sms.fromNumber || "",
  };
}

export async function sendTestSms({ phoneNumber, message }) {
  if (!isSmsConfigured()) {
    return {
      ok: false,
      code: 503,
      error: "HAL SMS is not configured yet. Add Twilio credentials in .env first.",
    };
  }

  if ((config.sms.provider || "").toLowerCase() !== "twilio") {
    return {
      ok: false,
      code: 400,
      error: "HAL only supports Twilio SMS in this scaffold right now.",
    };
  }

  const params = new URLSearchParams();
  params.set("To", phoneNumber);
  params.set("From", config.sms.fromNumber);
  params.set("Body", message);

  const auth = Buffer.from(`${config.sms.accountSid}:${config.sms.authToken}`).toString("base64");
  const response = await fetch(`https://api.twilio.com/2010-04-01/Accounts/${config.sms.accountSid}/Messages.json`, {
    method: "POST",
    headers: {
      Authorization: `Basic ${auth}`,
      "Content-Type": "application/x-www-form-urlencoded",
    },
    body: params,
  });

  const payload = await response.json();
  if (!response.ok) {
    return {
      ok: false,
      code: response.status,
      error: payload?.message || "Twilio rejected the SMS request.",
    };
  }

  return {
    ok: true,
    sid: payload.sid,
    status: payload.status,
    to: payload.to,
    from: payload.from,
  };
}

export function buildReminderSmsMessage(reminder) {
  return `HAL reminder: ${reminder.text} at ${reminder.at}`;
}
