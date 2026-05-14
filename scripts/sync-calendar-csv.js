import { randomUUID } from "node:crypto";
import { readFile, readdir, stat } from "node:fs/promises";
import path from "node:path";

import { config } from "../server/config.js";
import { readHalState, writeHalState } from "../server/state-store.js";

async function main() {
  const csvPath = await resolveCsvCalendarPath(config.csvCalendar.filePath);
  const csv = await readFile(csvPath, "utf8");
  const snapshot = await readHalState();
  const nextState = ensureHalState(snapshot.state);
  const csvFileName = path.basename(csvPath);

  const currentContent = nextState.calendarSettings?.uploadedCsvContent || "";
  const currentName = nextState.calendarSettings?.uploadedCsvName || "";
  if (currentContent === csv && currentName === csvFileName) {
    console.log(`HAL calendar sync skipped. No CSV change detected in ${csvPath}.`);
    return;
  }

  nextState.calendarSettings = {
    ...nextState.calendarSettings,
    provider: "csv",
    uploadedCsvContent: csv,
    uploadedCsvName: csvFileName,
    uploadedCsvImportedAt: new Date().toISOString(),
  };
  nextState._meta = {
    ...(nextState._meta || {}),
    savedAt: new Date().toISOString(),
    storageVersion: 1,
    syncMode: "local-server-mirror",
  };

  await writeHalState(nextState);

  const eventCount = countCsvEventRows(csv);
  console.log(`HAL calendar sync uploaded ${eventCount} CSV row(s) from ${csvPath}.`);
}

function ensureHalState(existingState) {
  if (existingState && typeof existingState === "object") {
    existingState.calendarSettings = {
      provider: "csv",
      googleMode: "both",
      googleWorkCalendarName: "Work",
      googlePersonalCalendarName: "Personal",
      uploadedCsvContent: "",
      uploadedCsvName: "",
      uploadedCsvImportedAt: "",
      ...(existingState.calendarSettings || {}),
    };
    existingState._meta = {
      savedAt: existingState._meta?.savedAt || null,
      storageVersion: 1,
      syncMode: existingState._meta?.syncMode || "local-server-mirror",
    };
    return existingState;
  }

  return {
    _meta: {
      savedAt: null,
      storageVersion: 1,
      syncMode: "local-server-mirror",
    },
    theme: "dark",
    voiceResponsesEnabled: true,
    quote: "Build the life and systems you want to live inside.",
    quickLinks: [],
    taskLists: [{ id: randomUUID(), name: "Regular Tasks" }],
    selectedTaskListId: "",
    showCompletedTasks: false,
    tasks: [],
    ideas: [],
    meetingNotes: [],
    calendarEvents: [],
    calendarSettings: {
      provider: "csv",
      googleMode: "both",
      googleWorkCalendarName: "Work",
      googlePersonalCalendarName: "Personal",
      uploadedCsvContent: "",
      uploadedCsvName: "",
      uploadedCsvImportedAt: "",
    },
    lastCalendarSyncDate: "",
    reminders: [],
    notificationSettings: {
      teamsEnabled: false,
      smsEnabled: false,
      phoneNumber: "",
    },
    followUps: [],
    audioInbox: [],
    selected: { type: null, id: null },
    lastCreated: { type: null, id: null },
    halMemory: {
      promptHistory: [],
      titleCorrections: [],
      intentCorrections: [],
      routeCorrections: [],
      routingPatterns: [],
      noteTargetCorrections: [],
    },
  };
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

function countCsvEventRows(csv) {
  return String(csv || "")
    .split(/\r?\n/)
    .slice(1)
    .map((line) => line.trim())
    .filter(Boolean)
    .length;
}

main().catch((error) => {
  console.error(`HAL calendar sync failed: ${error.message}`);
  process.exitCode = 1;
});
