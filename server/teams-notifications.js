import { config, isTeamsNotificationsConfigured } from "./config.js";

const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";

export function getTeamsNotificationStatus({ authenticated = false } = {}) {
  return {
    teamsEnabled: true,
    authenticated,
    providerConfigured: isTeamsNotificationsConfigured(),
    manifestId: config.teams.manifestId || "",
    detail: authenticated
      ? "HAL can send Teams activity notifications once the app is installed and the Graph scopes are consented."
      : "Sign in with Microsoft before HAL can send Teams activity notifications.",
    reason: !config.entra.clientId
      ? "HAL is missing Microsoft Entra configuration."
      : !config.teams.manifestId
        ? "Add HAL_TEAMS_MANIFEST_ID to .env to enable Teams reminder notifications."
        : "",
  };
}

export async function sendTestTeamsNotification(accessToken, { message }) {
  if (!isTeamsNotificationsConfigured()) {
    return {
      ok: false,
      code: 503,
      error: "HAL Teams notifications are not configured yet. Add the Teams manifest ID and Microsoft sign-in settings first.",
    };
  }

  const me = await graphFetch(accessToken, "/me?$select=id,displayName,userPrincipalName");
  const payload = {
    topic: {
      source: "text",
      value: "HAL Reminder",
      webUrl: config.teams.topicUrl || `${config.baseUrl}${config.tabPath}`,
    },
    activityType: "reminderTriggered",
    previewText: {
      content: message,
    },
    teamsAppId: config.teams.manifestId,
    templateParameters: [
      {
        name: "reminderText",
        value: message,
      },
    ],
  };

  const response = await fetch(`${GRAPH_BASE_URL}/users/${me.id}/teamwork/sendActivityNotification`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(payload),
  });

  if (!response.ok) {
    const errorPayload = await safeGraphError(response);
    return {
      ok: false,
      code: response.status,
      error: errorPayload || "Microsoft Graph rejected the Teams notification request.",
    };
  }

  return {
    ok: true,
    recipient: me.displayName || me.userPrincipalName || "your Teams account",
  };
}

async function graphFetch(accessToken, path) {
  const response = await fetch(`${GRAPH_BASE_URL}${path}`, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
    },
  });

  if (!response.ok) {
    const message = await safeGraphError(response);
    throw new Error(message || "Microsoft Graph request failed.");
  }

  return response.json();
}

async function safeGraphError(response) {
  try {
    const payload = await response.json();
    return payload?.error?.message || payload?.message || "";
  } catch {
    return "";
  }
}
