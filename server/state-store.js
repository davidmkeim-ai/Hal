import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { createClient } from "@supabase/supabase-js";
import { config, isSupabaseConfigured } from "./config.js";

const DATA_DIR = config.storage.dataRoot;
const STATE_FILE = path.resolve(DATA_DIR, "hal-state.json");
const SUPABASE_TABLE = "hal_state";
let supabaseClient = null;

export async function readHalState() {
  if (shouldUseSupabaseStore()) {
    return readSupabaseState();
  }

  return readFileState();
}

export async function writeHalState(state) {
  if (shouldUseSupabaseStore()) {
    return writeSupabaseState(state);
  }

  return writeFileState(state);
}

export function getHalStatePath() {
  return shouldUseSupabaseStore()
    ? `supabase:${SUPABASE_TABLE}/${config.supabase.stateKey}`
    : STATE_FILE;
}

export async function getHalStateStoreStatus() {
  const provider = shouldUseSupabaseStore() ? "supabase" : "file";
  try {
    const snapshot = await readHalState();
    return {
      ok: true,
      provider,
      path: snapshot.path,
      meta: snapshot.meta,
    };
  } catch (error) {
    return {
      ok: false,
      provider,
      path: getHalStatePath(),
      error: error.message,
    };
  }
}

function shouldUseSupabaseStore() {
  return config.storage.provider === "supabase" && isSupabaseConfigured();
}

async function readFileState() {
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

async function writeFileState(state) {
  await mkdir(DATA_DIR, { recursive: true });
  await writeFile(STATE_FILE, JSON.stringify(state, null, 2), "utf8");
  return readFileState();
}

async function readSupabaseState() {
  const client = getSupabaseClient();
  const { data, error } = await client
    .from(SUPABASE_TABLE)
    .select("state, updated_at")
    .eq("id", config.supabase.stateKey)
    .maybeSingle();

  if (error) {
    throw new Error(error.message);
  }

  const state = data?.state || null;
  return {
    state,
    meta: {
      ...buildStateMeta(state),
      savedAt: state?._meta?.savedAt || data?.updated_at || null,
      provider: "supabase",
    },
    path: getHalStatePath(),
  };
}

async function writeSupabaseState(state) {
  const client = getSupabaseClient();
  const { error } = await client
    .from(SUPABASE_TABLE)
    .upsert({
      id: config.supabase.stateKey,
      state,
      updated_at: new Date().toISOString(),
    });

  if (error) {
    throw new Error(error.message);
  }

  return readSupabaseState();
}

function getSupabaseClient() {
  if (!supabaseClient) {
    supabaseClient = createClient(config.supabase.url, config.supabase.serviceRoleKey, {
      auth: {
        persistSession: false,
        autoRefreshToken: false,
      },
    });
  }
  return supabaseClient;
}

function buildStateMeta(state) {
  return {
    savedAt: state?._meta?.savedAt || null,
    syncMode: state?._meta?.syncMode || "local-only",
    provider: shouldUseSupabaseStore() ? "supabase" : "file",
    counts: {
      tasks: Array.isArray(state?.tasks) ? state.tasks.length : 0,
      ideas: Array.isArray(state?.ideas) ? state.ideas.length : 0,
      meetingNotes: Array.isArray(state?.meetingNotes) ? state.meetingNotes.length : 0,
      reminders: Array.isArray(state?.reminders) ? state.reminders.length : 0,
    },
  };
}
