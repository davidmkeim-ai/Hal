import express from "express";
import session from "express-session";
import { randomUUID } from "node:crypto";
import { mkdir, readFile, readdir, stat, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { config, isAiConfigured, isAuthConfigured, isGoogleCalendarConfigured } from "./config.js";
import { createMsalClient, getAuthCodeUrlRequest, getTokenRequest } from "./msal.js";
import { createTeamsCalendarEvent, getOutlookCalendarEventsForDay } from "./graph.js";
import { createHalInterpretation } from "./openai.js";
import { getSmsStatus, sendTestSms } from "./sms.js";
import { changeHalAccessPassword, getHalAuthStatus, verifyHalAccessPassword } from "./auth-store.js";
import { getHalStateStoreStatus, readHalState, writeHalState } from "./state-store.js";
import { getTeamsNotificationStatus, sendTestTeamsNotification } from "./teams-notifications.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const workspaceRoot = path.resolve(__dirname, "..");
const backupsRoot = config.storage.backupsRoot;
const clientRootFiles = new Map([
  ["app.js", "app.js"],
  ["styles.css", "styles.css"],
  ["favicon.svg", "favicon.svg"],
  ["hal.config.example.js", "hal.config.example.js"],
]);

const app = express();
const msalClient = createMsalClient();

app.use(express.json({ limit: "10mb" }));
app.use(
  session({
    name: "hal.sid",
    secret: config.sessionSecret,
    resave: false,
    saveUninitialized: false,
    cookie: {
      httpOnly: true,
      sameSite: "lax",
      secure: config.baseUrl.startsWith("https://"),
      maxAge: 1000 * 60 * 60 * 8,
    },
  })
);

app.use(express.static(workspaceRoot));
registerStaticAssetRoutes(app);

app.get("/api/health", (_request, response) => {
  response.json({
    ok: true,
    app: "HAL",
    authConfigured: isAuthConfigured(),
    aiConfigured: isAiConfigured(),
    sms: getSmsStatus(),
  });
});

app.get("/api/auth/status", async (request, response) => {
  try {
    const status = await getHalAuthStatus({
      authenticated: Boolean(request.session.halAuthenticated),
    });
    response.json(status);
  } catch (error) {
    response.status(500).json({
      error: `HAL could not read its security status: ${error.message}`,
    });
  }
});

app.post("/api/auth/login", async (request, response) => {
  const password = String(request.body?.password || "");
  if (!password) {
    response.status(400).json({
      error: "Enter the HAL password first.",
    });
    return;
  }

  try {
    const valid = await verifyHalAccessPassword(password);
    if (!valid) {
      response.status(401).json({
        error: "That password did not unlock HAL.",
      });
      return;
    }

    request.session.halAuthenticated = true;
    const status = await getHalAuthStatus({ authenticated: true });
    response.json({
      ok: true,
      status,
    });
  } catch (error) {
    response.status(500).json({
      error: `HAL could not validate the password: ${error.message}`,
    });
  }
});

app.post("/api/auth/logout", (request, response) => {
  request.session.destroy((error) => {
    if (error) {
      response.status(500).json({
        error: "HAL could not end this session cleanly.",
      });
      return;
    }

    response.clearCookie("hal.sid");
    response.json({
      ok: true,
    });
  });
});

app.post("/api/auth/password", requireHalAuthentication, async (request, response) => {
  const currentPassword = String(request.body?.currentPassword || "");
  const newPassword = String(request.body?.newPassword || "");

  try {
    const result = await changeHalAccessPassword({ currentPassword, newPassword });
    if (!result.ok) {
      response.status(result.code || 400).json({
        error: result.error,
      });
      return;
    }

    response.json({
      ok: true,
      updatedAt: result.updatedAt,
    });
  } catch (error) {
    response.status(500).json({
      error: `HAL could not update the password: ${error.message}`,
    });
  }
});

app.use("/api", requireHalAuthentication);

app.get("/api/session", (request, response) => {
  response.json({
    appAuthenticated: Boolean(request.session.halAuthenticated),
    authenticated: Boolean(request.session.account),
    user: request.session.account
      ? {
          name: request.session.account.name,
          username: request.session.account.username,
        }
      : null,
  });
});

app.get("/api/state", async (_request, response) => {
  try {
    const snapshot = await readHalState();
    response.json({
      ok: true,
      state: snapshot.state,
      meta: snapshot.meta,
      path: snapshot.path,
    });
  } catch (error) {
    response.status(500).json({
      error: `HAL could not read the local state mirror: ${error.message}`,
    });
  }
});

app.get("/api/state/status", async (_request, response) => {
  const status = await getHalStateStoreStatus();
  response.status(status.ok ? 200 : 500).json(status);
});

app.put("/api/state", async (request, response) => {
  try {
    const payload = request.body;
    if (!payload || typeof payload !== "object") {
      response.status(400).json({
        error: "HAL needs state data before it can update the local mirror.",
      });
      return;
    }

    const snapshot = await writeHalState(payload);

    response.json({
      ok: true,
      path: snapshot.path,
      meta: snapshot.meta,
    });
  } catch (error) {
    response.status(500).json({
      error: `HAL could not update the local state mirror: ${error.message}`,
    });
  }
});

app.get("/api/backups", async (_request, response) => {
  try {
    const snapshots = await listBackupSnapshots();
    response.json({
      ok: true,
      snapshots,
    });
  } catch (error) {
    response.status(500).json({
      error: `HAL could not read local snapshots: ${error.message}`,
    });
  }
});

app.post("/api/backups/local", async (request, response) => {
  try {
    const payload = request.body;
    if (!payload || typeof payload !== "object") {
      response.status(400).json({
        error: "HAL needs backup data before it can create a local snapshot.",
      });
      return;
    }

    await mkdir(backupsRoot, { recursive: true });
    const filename = `hal-backup-${buildTimestampForFilename(new Date())}.json`;
    const filePath = path.resolve(backupsRoot, filename);
    await writeFile(filePath, JSON.stringify(payload, null, 2), "utf8");

    response.status(201).json({
      ok: true,
      file: filePath,
      name: filename,
    });
  } catch (error) {
    response.status(500).json({
      error: `HAL could not create a local snapshot: ${error.message}`,
    });
  }
});

app.get("/api/calendar/day", async (request, response) => {
  if (!request.session.accessToken) {
    response.status(401).json({
      error: "Sign in with Microsoft before syncing Outlook calendar events.",
    });
    return;
  }

  const date = String(request.query.date || "").trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) {
    response.status(400).json({
      error: "A calendar date in YYYY-MM-DD format is required.",
    });
    return;
  }

  const timeZone = String(request.query.timeZone || "America/Chicago");

  try {
    const events = await getOutlookCalendarEventsForDay(request.session.accessToken, date, timeZone);
    response.json({
      ok: true,
      date,
      events,
    });
  } catch (error) {
    response.status(502).json({
      error: error.message,
    });
  }
});

app.get("/api/calendar/csv/day", async (request, response) => {
  const date = String(request.query.date || "").trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) {
    response.status(400).json({
      error: "A calendar date in YYYY-MM-DD format is required.",
    });
    return;
  }

  try {
    const csvPath = await resolveCsvCalendarPath(config.csvCalendar.filePath);
    const csv = await readFile(csvPath, "utf8");
    const events = parseCsvCalendarEventsForDay(csv, date);
    response.json({
      ok: true,
      date,
      events,
      source: "csv",
      filePath: csvPath,
    });
  } catch (error) {
    response.status(404).json({
      error: `HAL could not read the calendar CSV at ${config.csvCalendar.filePath}.`,
    });
  }
});

app.get("/api/notifications/teams/status", (request, response) => {
  response.json(getTeamsNotificationStatus({
    authenticated: Boolean(request.session.accessToken),
  }));
});

app.get("/api/google-calendar/status", (_request, response) => {
  response.json({
    configured: isGoogleCalendarConfigured(),
    provider: "google",
  });
});

app.get("/api/notifications/status", (_request, response) => {
  response.json(getSmsStatus());
});

app.post("/api/notifications/test", async (request, response) => {
  const validation = validateNotificationTestRequest(request.body);
  if (!validation.ok) {
    response.status(400).json({
      error: validation.error,
    });
    return;
  }

  try {
    const result = await sendTestSms(validation.payload);
    if (!result.ok) {
      response.status(result.code || 502).json({
        error: result.error,
      });
      return;
    }

    response.status(201).json({
      ok: true,
      message: `HAL sent a test text to ${result.to}.`,
      sid: result.sid,
      status: result.status,
    });
  } catch (error) {
    response.status(502).json({
      error: error.message,
    });
  }
});

app.post("/api/notifications/teams/test", async (request, response) => {
  if (!request.session.accessToken) {
    response.status(401).json({
      error: "Sign in with Microsoft before testing Teams notifications.",
    });
    return;
  }

  const message = String(request.body?.message || "HAL Teams notification test.").trim();
  if (!message) {
    response.status(400).json({
      error: "A Teams notification message is required.",
    });
    return;
  }

  try {
    const result = await sendTestTeamsNotification(request.session.accessToken, { message });
    if (!result.ok) {
      response.status(result.code || 502).json({
        error: result.error,
      });
      return;
    }

    response.status(201).json({
      ok: true,
      message: `HAL sent a Teams notification to ${result.recipient}.`,
    });
  } catch (error) {
    response.status(502).json({
      error: error.message,
    });
  }
});

app.use("/auth", requireHalAuthentication);

app.get("/auth/microsoft/start", async (request, response) => {
  if (!isAuthConfigured()) {
    response.status(500).send("HAL is missing Microsoft Entra configuration. Update .env first.");
    return;
  }

  if (!msalClient) {
    response.status(500).send("HAL could not initialize the Microsoft auth client.");
    return;
  }

  const stateId = randomUUID();
  request.session.authState = stateId;

  const authUrl = await msalClient.getAuthCodeUrl({
    ...getAuthCodeUrlRequest(),
    state: stateId,
  });
  response.redirect(authUrl);
});

app.get("/auth/microsoft/callback", async (request, response) => {
  try {
    if (!msalClient) {
      response.status(500).send("HAL could not initialize the Microsoft auth client.");
      return;
    }

    if (!request.query.code || !request.query.state) {
      response.status(400).send("HAL did not receive the expected Microsoft auth response.");
      return;
    }

    if (request.query.state !== request.session.authState) {
      response.status(400).send("HAL rejected the sign-in callback because the auth state did not match.");
      return;
    }

    const tokenResponse = await msalClient.acquireTokenByCode({
      ...getTokenRequest(String(request.query.code)),
      state: String(request.query.state),
    });

    request.session.account = tokenResponse.account;
    request.session.accessToken = tokenResponse.accessToken;
    request.session.idTokenClaims = tokenResponse.idTokenClaims || null;
    request.session.authState = null;

    response.redirect(config.tabPath);
  } catch (error) {
    response.status(500).send(`HAL could not finish Microsoft sign-in: ${error.message}`);
  }
});

app.post("/api/meetings", async (request, response) => {
  if (!request.session.accessToken) {
    response.status(401).json({
      error: "Sign in with Microsoft before creating a Teams meeting.",
    });
    return;
  }

  const validation = validateMeetingRequest(request.body);
  if (!validation.ok) {
    response.status(400).json({
      error: validation.error,
    });
    return;
  }

  try {
    const event = await createTeamsCalendarEvent(request.session.accessToken, validation.meeting);
    response.status(201).json({ event });
  } catch (error) {
    response.status(502).json({
      error: error.message,
    });
  }
});

app.post("/api/ai/interpret", async (request, response) => {
  if (!isAiConfigured()) {
    response.json({
      enabled: false,
      reason: "HAL AI is not configured yet.",
    });
    return;
  }

  const validation = validateInterpretationRequest(request.body);
  if (!validation.ok) {
    response.status(400).json({
      enabled: true,
      error: validation.error,
    });
    return;
  }

  try {
    const interpretation = await createHalInterpretation(validation.payload);
    response.json({
      enabled: true,
      interpretation,
    });
  } catch (error) {
    response.status(502).json({
      enabled: true,
      error: error.message,
    });
  }
});

app.get("*", (_request, response) => {
  response.sendFile(path.join(workspaceRoot, "index.html"));
});

export default app;

if (isDirectExecution()) {
  app.listen(config.port, () => {
    console.log(`HAL server listening on ${config.baseUrl}`);
  });
}

function isDirectExecution() {
  return path.resolve(process.argv[1] || "") === __filename;
}

function registerStaticAssetRoutes(expressApp) {
  clientRootFiles.forEach((filePath, routePath) => {
    expressApp.get(`/${routePath}`, (_request, response) => {
      response.sendFile(path.join(workspaceRoot, filePath));
    });
  });

  expressApp.get("/assets/*", (request, response) => {
    const assetPath = String(request.params[0] || "").replace(/\\/g, "/");
    if (!assetPath || assetPath.includes("..")) {
      response.status(404).end();
      return;
    }

    response.sendFile(path.join(workspaceRoot, "assets", assetPath));
  });
}

function requireHalAuthentication(request, response, next) {
  if (request.session?.halAuthenticated) {
    next();
    return;
  }

  response.status(401).json({
    error: "HAL is locked. Sign in first.",
  });
}

async function listBackupSnapshots() {
  await mkdir(backupsRoot, { recursive: true });
  const entries = await readdir(backupsRoot);
  const snapshots = await Promise.all(entries
    .filter((entry) => entry.toLowerCase().endsWith(".json"))
    .map(async (entry) => {
      const fullPath = path.resolve(backupsRoot, entry);
      const details = await stat(fullPath);
      return {
        name: entry,
        path: fullPath,
        modifiedAt: details.mtime.toISOString(),
        modifiedLabel: details.mtime.toLocaleString("en-US"),
      };
    }));

  return snapshots
    .sort((left, right) => new Date(right.modifiedAt) - new Date(left.modifiedAt))
    .slice(0, 10);
}

function buildTimestampForFilename(date) {
  const year = String(date.getFullYear());
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  const hours = String(date.getHours()).padStart(2, "0");
  const minutes = String(date.getMinutes()).padStart(2, "0");
  const seconds = String(date.getSeconds()).padStart(2, "0");
  return `${year}${month}${day}-${hours}${minutes}${seconds}`;
}

function parseCsvCalendarEventsForDay(csv, selectedDate) {
  const lines = String(csv || "")
    .replace(/^\uFEFF/, "")
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);

  if (!lines.length) {
    return [];
  }

  const rows = lines.slice(1);
  return rows
    .map(parseCsvCalendarRow)
    .filter(Boolean)
    .filter((event) => event.date === selectedDate)
    .sort((left, right) => compareEventTimes(left.time, right.time));
}

async function resolveCsvCalendarPath(inputPath) {
  const targetPath = path.resolve(inputPath);
  const targetStats = await stat(targetPath);
  if (targetStats.isFile()) {
    return targetPath;
  }

  if (!targetStats.isDirectory()) {
    throw new Error("CSV calendar path is not a file or directory.");
  }

  const nestedSameName = path.join(targetPath, path.basename(targetPath));
  try {
    const nestedStats = await stat(nestedSameName);
    if (nestedStats.isFile()) {
      return nestedSameName;
    }
  } catch {
    // Fall back to scanning the directory.
  }

  const entries = await readdir(targetPath, { withFileTypes: true });
  const firstCsvFile = entries.find((entry) => entry.isFile() && entry.name.toLowerCase().endsWith(".csv"));
  if (firstCsvFile) {
    return path.join(targetPath, firstCsvFile.name);
  }

  throw new Error("No CSV file was found in the calendar export directory.");
}

function parseCsvCalendarRow(line) {
  const values = splitSimpleCsvLine(line);
  if (values.length < 3) {
    return null;
  }

  const title = values[0]?.trim();
  const date = normalizeCsvDate(values[1]);
  const time = normalizeCsvTime(values[2]);

  if (!title || !date) {
    return null;
  }

  return {
    id: randomUUID(),
    title,
    date,
    time,
    source: "csv",
  };
}

function splitSimpleCsvLine(line) {
  const values = [];
  let current = "";
  let inQuotes = false;

  for (let index = 0; index < line.length; index += 1) {
    const char = line[index];
    const nextChar = line[index + 1];

    if (char === '"') {
      if (inQuotes && nextChar === '"') {
        current += '"';
        index += 1;
        continue;
      }
      inQuotes = !inQuotes;
      continue;
    }

    if (char === "," && !inQuotes) {
      values.push(current);
      current = "";
      continue;
    }

    current += char;
  }

  values.push(current);
  return values;
}

function normalizeCsvDate(value) {
  const raw = String(value || "").trim();
  if (!raw) {
    return "";
  }

  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) {
    return raw;
  }

  const slashMatch = raw.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (slashMatch) {
    const month = slashMatch[1].padStart(2, "0");
    const day = slashMatch[2].padStart(2, "0");
    return `${slashMatch[3]}-${month}-${day}`;
  }

  const parsed = new Date(raw);
  if (Number.isNaN(parsed.getTime())) {
    return "";
  }

  const month = String(parsed.getMonth() + 1).padStart(2, "0");
  const day = String(parsed.getDate()).padStart(2, "0");
  const year = parsed.getFullYear();
  return `${year}-${month}-${day}`;
}

function normalizeCsvTime(value) {
  return String(value || "").trim();
}

function compareEventTimes(left, right) {
  const leftMinutes = parseTimeToMinutes(left);
  const rightMinutes = parseTimeToMinutes(right);
  return leftMinutes - rightMinutes;
}

function parseTimeToMinutes(value) {
  const raw = String(value || "").trim().toLowerCase();
  if (!raw) {
    return Number.MAX_SAFE_INTEGER;
  }

  const match = raw.match(/^(\d{1,2})(?::(\d{2}))?\s*(am|pm)?$/);
  if (!match) {
    return Number.MAX_SAFE_INTEGER - 1;
  }

  let hours = Number(match[1]);
  const minutes = Number(match[2] || "0");
  const meridiem = match[3];

  if (meridiem === "pm" && hours < 12) {
    hours += 12;
  }
  if (meridiem === "am" && hours === 12) {
    hours = 0;
  }

  return (hours * 60) + minutes;
}

function validateMeetingRequest(body) {
  if (!body?.title?.trim()) {
    return { ok: false, error: "Meeting title is required." };
  }

  if (!body?.start?.trim()) {
    return { ok: false, error: "Meeting start time is required." };
  }

  if (!body?.timeZone?.trim()) {
    return { ok: false, error: "A valid time zone is required." };
  }

  if (!body?.end?.trim()) {
    return { ok: false, error: "Meeting end time is required." };
  }

  const startDate = new Date(body.start);
  if (Number.isNaN(startDate.getTime())) {
    return { ok: false, error: "Start time must be a valid ISO timestamp." };
  }

  const endDate = new Date(body.end);
  if (Number.isNaN(endDate.getTime()) || endDate <= startDate) {
    return { ok: false, error: "End time must be later than the start time." };
  }

  return {
    ok: true,
    meeting: {
      title: body.title.trim(),
      notes: String(body.notes || ""),
      category: String(body.category || ""),
      start: normalizeUtcDateTime(body.start),
      end: normalizeUtcDateTime(body.end),
      timeZone: body.timeZone.trim(),
      sourceTimeZone: String(body.sourceTimeZone || ""),
      attendees: Array.isArray(body.attendees) ? body.attendees : [],
    },
  };
}

function normalizeUtcDateTime(value) {
  const date = new Date(value);
  const year = date.getUTCFullYear();
  const month = String(date.getUTCMonth() + 1).padStart(2, "0");
  const day = String(date.getUTCDate()).padStart(2, "0");
  const hours = String(date.getUTCHours()).padStart(2, "0");
  const minutes = String(date.getUTCMinutes()).padStart(2, "0");
  const seconds = String(date.getUTCSeconds()).padStart(2, "0");
  return `${year}-${month}-${day}T${hours}:${minutes}:${seconds}`;
}

function validateInterpretationRequest(body) {
  if (!body?.text || typeof body.text !== "string") {
    return { ok: false, error: "Interpretation text is required." };
  }

  return {
    ok: true,
    payload: {
      text: body.text,
      today: String(body.today || ""),
      timeZone: String(body.timeZone || ""),
      taskLists: Array.isArray(body.taskLists) ? body.taskLists : [],
      localInterpretation: sanitizeInterpretationSummary(body.localInterpretation),
      memory: sanitizeMemory(body.memory),
    },
  };
}

function validateNotificationTestRequest(body) {
  if (!body?.phoneNumber || typeof body.phoneNumber !== "string") {
    return { ok: false, error: "A destination phone number is required." };
  }

  const phoneNumber = body.phoneNumber.trim();
  if (!/^\+\d{10,15}$/.test(phoneNumber)) {
    return { ok: false, error: "Phone number must be in E.164 format, like +15551234567." };
  }

  return {
    ok: true,
    payload: {
      phoneNumber,
      message: String(body.message || "HAL test message."),
    },
  };
}

function sanitizeInterpretationSummary(interpretation) {
  if (!interpretation || typeof interpretation !== "object") {
    return null;
  }

  return {
    type: String(interpretation.type || ""),
    title: String(interpretation.title || ""),
    label: String(interpretation.label || ""),
    preview: String(interpretation.preview || ""),
    dueDate: String(interpretation.dueDate || ""),
    dueTime: String(interpretation.dueTime || ""),
    person: String(interpretation.person || ""),
  };
}

function sanitizeMemory(memory) {
  if (!memory || typeof memory !== "object") {
    return {
      recentPrompts: [],
      corrections: [],
      preferences: {
        taskTitleCorrections: [],
        ideaTitleCorrections: [],
        intentCorrections: [],
        taskListCorrections: [],
        meetingTargetCorrections: [],
        routingPatterns: [],
      },
    };
  }

  return {
    recentPrompts: Array.isArray(memory.recentPrompts) ? memory.recentPrompts.slice(0, 8) : [],
    corrections: Array.isArray(memory.corrections) ? memory.corrections.slice(0, 12) : [],
    preferences: {
      taskTitleCorrections: Array.isArray(memory.preferences?.taskTitleCorrections)
        ? memory.preferences.taskTitleCorrections.slice(0, 12)
        : [],
      ideaTitleCorrections: Array.isArray(memory.preferences?.ideaTitleCorrections)
        ? memory.preferences.ideaTitleCorrections.slice(0, 12)
        : [],
      intentCorrections: Array.isArray(memory.preferences?.intentCorrections)
        ? memory.preferences.intentCorrections.slice(0, 12)
        : [],
      taskListCorrections: Array.isArray(memory.preferences?.taskListCorrections)
        ? memory.preferences.taskListCorrections.slice(0, 12)
        : [],
      meetingTargetCorrections: Array.isArray(memory.preferences?.meetingTargetCorrections)
        ? memory.preferences.meetingTargetCorrections.slice(0, 12)
        : [],
      routingPatterns: Array.isArray(memory.preferences?.routingPatterns)
        ? memory.preferences.routingPatterns.slice(0, 16)
        : [],
    },
  };
}
