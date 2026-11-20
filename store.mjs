import { cloudMode, readCloud, saveCloud, loadCloud } from './cloud.mjs';
export const PREF_KEY = 'referral-office-preferences-v1';
export function readData(){return readCloud();}
export function persist(data,options){if(!cloudMode)throw Error('Configure the workspace before saving.');return saveCloud(data,options);}
export async function refreshData(){if(!cloudMode)throw Error('Configure the workspace before loading records.');return loadCloud();}
export function readPreferences() { try { const value = JSON.parse(localStorage.getItem(PREF_KEY)); return value && typeof value === 'object' && !Array.isArray(value) ? value : {}; } catch { return {}; } }
export function savePreferences(value) { localStorage.setItem(PREF_KEY, JSON.stringify(value)); }
export function changeEvents(previous, next, type, actor, now = new Date().toISOString()) {
  const ignored = new Set(['key', 'version', 'updatedAt', 'contacts']);
  return Object.keys(next).filter(key => !ignored.has(key) && JSON.stringify(previous?.[key] ?? '') !== JSON.stringify(next[key] ?? '')).map(field => ({ key: crypto.randomUUID(), recordKey: next.key, recordType: type, field, before: previous?.[field] ?? '', after: next[field] ?? '', actor, at: now }));
}

