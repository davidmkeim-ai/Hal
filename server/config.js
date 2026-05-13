import dotenv from "dotenv";
import path from "node:path";

dotenv.config();

const scopesFromEnv = process.env.GRAPH_SCOPES
  ? process.env.GRAPH_SCOPES.split(/\s+/).filter(Boolean)
  : ["User.Read", "Calendars.ReadWrite", "TeamsActivity.Send", "offline_access", "openid", "profile"];

export const config = {
  port: Number(process.env.PORT || 4000),
  sessionSecret: process.env.SESSION_SECRET || "hal-local-dev-secret",
  baseUrl: (process.env.HAL_BASE_URL || "http://localhost:4000").replace(/\/$/, ""),
  tabPath: process.env.HAL_TAB_PATH || "/index.html",
  storage: {
    provider: (process.env.HAL_STORAGE_PROVIDER || "file").toLowerCase(),
    dataRoot: process.env.HAL_DATA_ROOT
      ? path.resolve(process.env.HAL_DATA_ROOT)
      : path.resolve(process.cwd(), "data"),
    backupsRoot: process.env.HAL_BACKUP_ROOT
      ? path.resolve(process.env.HAL_BACKUP_ROOT)
      : path.resolve(process.cwd(), "backups"),
  },
  supabase: {
    url: process.env.SUPABASE_URL || "",
    serviceRoleKey: process.env.SUPABASE_SERVICE_ROLE_KEY || "",
    stateKey: process.env.HAL_SUPABASE_STATE_KEY || "default",
  },
  openai: {
    apiKey: process.env.OPENAI_API_KEY || "",
    model: process.env.HAL_AI_MODEL || "gpt-5-mini",
  },
  sms: {
    provider: process.env.HAL_SMS_PROVIDER || "",
    fromNumber: process.env.HAL_SMS_FROM_NUMBER || "",
    accountSid: process.env.TWILIO_ACCOUNT_SID || "",
    authToken: process.env.TWILIO_AUTH_TOKEN || "",
  },
  teams: {
    manifestId: process.env.HAL_TEAMS_MANIFEST_ID || "",
    topicUrl: process.env.HAL_TEAMS_TOPIC_URL || "",
  },
  google: {
    clientId: process.env.GOOGLE_CLIENT_ID || "",
    clientSecret: process.env.GOOGLE_CLIENT_SECRET || "",
    redirectUri: process.env.GOOGLE_REDIRECT_URI || "",
  },
  csvCalendar: {
    filePath: process.env.HAL_CSV_CALENDAR_PATH
      ? path.resolve(process.env.HAL_CSV_CALENDAR_PATH)
      : path.resolve(process.cwd(), "calendar_export.csv"),
  },
  entra: {
    clientId: process.env.ENTRA_CLIENT_ID || "",
    clientSecret: process.env.ENTRA_CLIENT_SECRET || "",
    tenantId: process.env.ENTRA_TENANT_ID || "common",
    scopes: scopesFromEnv,
  },
};

export function getRedirectUri() {
  return `${config.baseUrl}/auth/microsoft/callback`;
}

export function isAuthConfigured() {
  return Boolean(config.entra.clientId && config.entra.clientSecret);
}

export function isAiConfigured() {
  return Boolean(config.openai.apiKey);
}

export function isSmsConfigured() {
  if ((config.sms.provider || "").toLowerCase() !== "twilio") {
    return false;
  }

  return Boolean(config.sms.fromNumber && config.sms.accountSid && config.sms.authToken);
}

export function isTeamsNotificationsConfigured() {
  return Boolean(config.entra.clientId && config.teams.manifestId);
}

export function isGoogleCalendarConfigured() {
  return Boolean(config.google.clientId && config.google.clientSecret && config.google.redirectUri);
}

export function isSupabaseConfigured() {
  return Boolean(config.supabase.url && config.supabase.serviceRoleKey);
}
