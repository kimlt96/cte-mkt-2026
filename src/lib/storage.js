/**
 * Persistent storage layer for the standalone app.
 *
 * The original prototype (built as a Claude.ai artifact) used the
 * `window.storage` key-value API provided by that sandbox. This app runs
 * as a normal website, so we implement the same get/set/delete/list
 * contract on top of IndexedDB (via idb-keyval), which — per the original
 * spec — is preferred over localStorage because the app stores many
 * base64 report images.
 *
 * Swapping this file for a real backend (e.g. Supabase/PostgreSQL) later
 * only requires re-implementing these four functions; nothing in the UI
 * layer needs to change.
 */
import { get as idbGet, set as idbSet, del as idbDel, keys as idbKeys } from "idb-keyval";

export async function storageGet(key) {
  try {
    const raw = await idbGet(key);
    if (raw === undefined || raw === null) return null;
    return { key, value: raw, shared: false };
  } catch (e) {
    console.error("storage.get failed", key, e);
    return null;
  }
}

export async function storageSet(key, value) {
  try {
    await idbSet(key, value);
    return { key, value, shared: false };
  } catch (e) {
    console.error("storage.set failed", key, e);
    return null;
  }
}

export async function storageDelete(key) {
  try {
    await idbDel(key);
    return { key, deleted: true, shared: false };
  } catch (e) {
    console.error("storage.delete failed", key, e);
    return null;
  }
}

export async function storageList(prefix = "") {
  try {
    const all = await idbKeys();
    const keys = all.filter((k) => typeof k === "string" && k.startsWith(prefix));
    return { keys, prefix, shared: false };
  } catch (e) {
    console.error("storage.list failed", prefix, e);
    return null;
  }
}
