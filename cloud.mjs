let client, session, snapshot, config, authEpoch = 0;
export let cloudMode = false;
export let member = null;
export const emptyData = () => ({ version: 1, revision: 0, patients: [], requests: [], audit: [], contactOutcomes: [], imports: [], staff: [], settings: { necessary: ['phone','centre','building','apartment'], hospitals: [], agencies: [] } });
export async function initialiseCloud() {
  config = await fetch('./config.json', { cache: 'no-store' }).then(r => r.ok ? r.json() : {}).catch(() => ({}));
  cloudMode = !!config.supabaseUrl;
  if (cloudMode) {
    validateConfig(config);
    const { createClient } = await import('./vendor/supabase.mjs');
    client = createClient(config.supabaseUrl, config.publishableKey, { auth: { persistSession: true, autoRefreshToken: true, detectSessionInUrl: false } });
    client.auth.onAuthStateChange((event, value) => { session = value; if(event === 'SIGNED_OUT'){ authEpoch++; snapshot=null; member=null; } });
  }
}
export async function restoreSignIn() {
  const { data, error } = await client.auth.getSession();
  if (error) throw Error(error.message);
  session = data.session;
  // Access and staff membership are checked by the database on every reload.
  // Only the sign-in session is persisted; workspace records stay in memory.
  if (session) await loadCloud();
}
export function validateConfig(value) {
  const url = new URL(value.supabaseUrl);
  if (url.protocol !== 'https:' || !/^[a-z0-9-]+\.supabase\.co$/.test(url.hostname) || url.pathname !== '/' || url.search || url.hash || url.username || url.password) throw Error('Enter your Supabase project URL: https://PROJECT.supabase.co');
  const key = value.publishableKey || '';
  if (!key.startsWith('sb_publishable_')) {
    try { const payload = JSON.parse(atob(key.split('.')[1].replace(/-/g,'+').replace(/_/g,'/'))); if (payload.role !== 'anon') throw Error(); }
    catch { throw Error('Use the publishable key or legacy anon key. Never use a secret or service-role key.'); }
  }
  return value;
}
export async function signIn(email, password) {
  const result = await client.auth.signInWithPassword({ email, password });
  if (result.error) throw Error('Sign-in failed. Check your email and password.');
  session = result.data.session;
  try { return await loadCloud(); } catch (error) { await signOut(); throw error; }
}
export async function signOut() {
  authEpoch++;
  session = null; member = null; snapshot = null;
  const result = await client?.auth.signOut({ scope: 'local' });
  return result;
}
export function readCloud() { return structuredClone(snapshot || emptyData()); }
export async function loadCloud() {
  if (!session) throw Error('Sign in to load records.');
  const epoch = authEpoch;
  const { data, error } = await client.rpc('erc_load');
  if (error) throw Error(error.message);
  if(epoch !== authEpoch || !session) throw Error('Sign-in changed. Sign in again to continue.');
  member = data.member; snapshot = data; return readCloud();
}
export function changesSince(previous, next) {
  const changes = {};
  for (const key of ['patients','requests','imports','contactOutcomes']) {
    const lookup = new Map((previous[key] || []).map(row => [row.key, row]));
    changes[key] = (next[key] || []).filter(row => JSON.stringify(row) !== JSON.stringify(lookup.get(row.key)));
  }
  if (JSON.stringify(previous.settings) !== JSON.stringify(next.settings)) changes.settings = next.settings;
  return changes;
}
export async function saveCloud(next, { reason = '' } = {}) {
  if (!session || !snapshot) throw Error('Sign in before saving.');
  const epoch = authEpoch;
  const { data, error } = await client.rpc('erc_commit', { expected_revision: next.revision - 1, changes: changesSince(snapshot, next), reason });
  if (error) {
    const failure = Error(error.message);
    failure.conflict = error.code === '40001' || error.message.includes('ERC_CONFLICT');
    throw failure;
  }
  if(epoch !== authEpoch || !session) throw Error('Sign-in changed. Sign in again to continue.');
  snapshot = data; member = data.member; return readCloud();
}
export async function restoreCloud(backup) {
  if (!session || member?.role !== 'admin') throw Error('Administrator sign-in is required.');
  const current = await loadCloud();
  const epoch = authEpoch;
  const { data, error } = await client.rpc('erc_restore', { expected_revision: current.revision, backup });
  if (error) throw Error(error.message);
  if(epoch !== authEpoch || !session) throw Error('Sign-in changed. Sign in again to continue.');
  snapshot = data; member = data.member; return readCloud();
}
