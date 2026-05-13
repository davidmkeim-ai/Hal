import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { randomBytes, scrypt as scryptCallback, timingSafeEqual } from "node:crypto";
import { promisify } from "node:util";
import { createClient } from "@supabase/supabase-js";
import { config, isPasswordBootstrapConfigured, isSupabaseConfigured } from "./config.js";

const scrypt = promisify(scryptCallback);
const DATA_DIR = config.storage.dataRoot;
const AUTH_FILE = path.resolve(DATA_DIR, "hal-auth.json");
const SUPABASE_TABLE = "hal_state";
let supabaseClient = null;

export async function getHalAuthStatus({ authenticated = false } = {}) {
  const settings = await readAuthSettings();
  const hasStoredPassword = Boolean(settings?.passwordHash);
  return {
    authenticated,
    passwordConfigured: hasStoredPassword || isPasswordBootstrapConfigured(),
    hasStoredPassword,
    usingBootstrapPassword: !hasStoredPassword && isPasswordBootstrapConfigured(),
    provider: shouldUseSupabaseStore() ? "supabase" : "file",
  };
}

export async function verifyHalAccessPassword(password) {
  const trimmed = String(password || "");
  if (!trimmed) {
    return false;
  }

  const settings = await readAuthSettings();
  if (settings?.passwordHash) {
    return verifyPasswordHash(trimmed, settings.passwordHash);
  }

  if (!isPasswordBootstrapConfigured()) {
    return false;
  }

  return trimmed === config.auth.bootstrapPassword;
}

export async function changeHalAccessPassword({ currentPassword, newPassword }) {
  const validCurrent = await verifyHalAccessPassword(currentPassword);
  if (!validCurrent) {
    return {
      ok: false,
      code: 401,
      error: "Current password is incorrect.",
    };
  }

  const trimmedNew = String(newPassword || "");
  if (trimmedNew.length < 8) {
    return {
      ok: false,
      code: 400,
      error: "Use a password with at least 8 characters.",
    };
  }

  const settings = {
    passwordHash: await createPasswordHash(trimmedNew),
    updatedAt: new Date().toISOString(),
  };
  await writeAuthSettings(settings);

  return {
    ok: true,
    updatedAt: settings.updatedAt,
  };
}

async function readAuthSettings() {
  if (shouldUseSupabaseStore()) {
    return readSupabaseAuthSettings();
  }

  return readFileAuthSettings();
}

async function writeAuthSettings(settings) {
  if (shouldUseSupabaseStore()) {
    return writeSupabaseAuthSettings(settings);
  }

  return writeFileAuthSettings(settings);
}

function shouldUseSupabaseStore() {
  return config.storage.provider === "supabase" && isSupabaseConfigured();
}

async function readFileAuthSettings() {
  try {
    await mkdir(DATA_DIR, { recursive: true });
    const raw = await readFile(AUTH_FILE, "utf8");
    return JSON.parse(raw);
  } catch (error) {
    if (error?.code === "ENOENT") {
      return null;
    }
    throw error;
  }
}

async function writeFileAuthSettings(settings) {
  await mkdir(DATA_DIR, { recursive: true });
  await writeFile(AUTH_FILE, JSON.stringify(settings, null, 2), "utf8");
  return settings;
}

async function readSupabaseAuthSettings() {
  const client = getSupabaseClient();
  const { data, error } = await client
    .from(SUPABASE_TABLE)
    .select("state")
    .eq("id", config.auth.stateKey)
    .maybeSingle();

  if (error) {
    throw new Error(error.message);
  }

  return data?.state || null;
}

async function writeSupabaseAuthSettings(settings) {
  const client = getSupabaseClient();
  const { error } = await client
    .from(SUPABASE_TABLE)
    .upsert({
      id: config.auth.stateKey,
      state: settings,
      updated_at: new Date().toISOString(),
    });

  if (error) {
    throw new Error(error.message);
  }

  return settings;
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

async function createPasswordHash(password) {
  const salt = randomBytes(16).toString("hex");
  const derived = await scrypt(password, salt, 64);
  return `${salt}:${Buffer.from(derived).toString("hex")}`;
}

async function verifyPasswordHash(password, hash) {
  const [salt, storedHex] = String(hash || "").split(":");
  if (!salt || !storedHex) {
    return false;
  }

  const derived = await scrypt(password, salt, 64);
  const stored = Buffer.from(storedHex, "hex");
  const candidate = Buffer.from(derived);

  if (stored.length !== candidate.length) {
    return false;
  }

  return timingSafeEqual(stored, candidate);
}
