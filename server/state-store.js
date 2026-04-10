import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { config } from "./config.js";

const DATA_DIR = config.storage.dataRoot;
const STATE_FILE = path.resolve(DATA_DIR, "hal-state.json");

export async function readHalState() {
  try {
    await mkdir(DATA_DIR, { recursive: true });
    const raw = await readFile(STATE_FILE, "utf8");
    const parsed = JSON.parse(raw);
    return {
      state: parsed,
      meta: buildStateMeta(parsed),
      path: STATE_FILE,
    };
  } catch (error) {
    if (error?.code === "ENOENT") {
      return {
        state: null,
        meta: buildStateMeta(null),
        path: STATE_FILE,
      };
    }
    throw error;
  }
}

export async function writeHalState(state) {
  await mkdir(DATA_DIR, { recursive: true });
  await writeFile(STATE_FILE, JSON.stringify(state, null, 2), "utf8");
  return readHalState();
}

export function getHalStatePath() {
  return STATE_FILE;
}

function buildStateMeta(state) {
  return {
    savedAt: state?._meta?.savedAt || null,
    syncMode: state?._meta?.syncMode || "local-only",
    counts: {
      tasks: Array.isArray(state?.tasks) ? state.tasks.length : 0,
      ideas: Array.isArray(state?.ideas) ? state.ideas.length : 0,
      meetingNotes: Array.isArray(state?.meetingNotes) ? state.meetingNotes.length : 0,
      reminders: Array.isArray(state?.reminders) ? state.reminders.length : 0,
    },
  };
}
