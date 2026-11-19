// Microsoft sign-in: OAuth 2.0 authorization code flow with PKCE for a single-page app.
// No client secret is involved; tokens stay in this browser.
import { CONFIG } from './config.js';
import { store } from './lib.js';

const TOKENS = 'ma.tokens';
const PKCE = 'ma.pkce';

export function msConfig() {
  const s = store.get('settings', {});
  return {
    clientId: (s.clientId || CONFIG.clientId || '').trim(),
    tenantId: (s.tenantId || CONFIG.tenantId || 'organizations').trim(),
  };
}

// Permissions asked for at sign-in: the read-only base plus the features switched on in Settings.
export function requestedScopes() {
  const s = store.get('settings', {});
  const x = CONFIG.extraScopes;
  const list = [...CONFIG.baseScopes];
  if (s.accSend) list.push(...x.send);
  if (s.accShared) list.push(...(s.accSend ? x.sharedSend : x.sharedRead));
  if (s.accContacts) list.push(...x.contacts);
  return [...new Set(list)];
}

// Permissions Microsoft actually granted, lower case without the Graph prefix, e.g. 'mail.read'.
const norm = (scope) => String(scope || '').split(/\s+/).filter(Boolean).map((x) => x.replace(/^.*\//, '').toLowerCase());
export function grantedScopes() {
  const t = getTokens();
  return new Set(t?.granted || norm((t?.scopes || []).join(' ')));
}

// The exact address IT registers as the SPA redirect URI, e.g. https://name.github.io/mail-assistant/
export const redirectUri = () => new URL('.', location.href).href;

const authority = (tenant = msConfig().tenantId) => `https://login.microsoftonline.com/${encodeURIComponent(tenant)}/oauth2/v2.0`;

const b64url = (buf) => btoa(String.fromCharCode(...new Uint8Array(buf))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
const randomString = (n = 48) => b64url(crypto.getRandomValues(new Uint8Array(n)));
const sha256 = (s) => crypto.subtle.digest('SHA-256', new TextEncoder().encode(s));

export async function login() {
  const { clientId, tenantId } = msConfig();
  if (!clientId) throw new Error('Add the Application (client) ID from IT first.');
  const verifier = randomString(48);
  const state = randomString(16);
  const scopes = requestedScopes();
  sessionStorage.setItem(PKCE, JSON.stringify({ verifier, state, clientId, tenantId, scopes }));
  const params = new URLSearchParams({
    client_id: clientId,
    response_type: 'code',
    redirect_uri: redirectUri(),
    response_mode: 'fragment',
    scope: scopes.join(' '),
    state,
    code_challenge: b64url(await sha256(verifier)),
    code_challenge_method: 'S256',
  });
  const hint = store.get('loginHint');
  if (hint) params.set('login_hint', hint);
  location.assign(`${authority(tenantId)}/authorize?${params}`);
}

// Called on page load: finishes sign-in if Microsoft just redirected back here.
export async function handleRedirect() {
  const hash = location.hash.slice(1);
  if (!/(^|&)(code|error)=/.test(hash)) return null;
  const params = new URLSearchParams(hash);
  history.replaceState(null, '', location.pathname + location.search);
  const saved = JSON.parse(sessionStorage.getItem(PKCE) || 'null');
  sessionStorage.removeItem(PKCE);
  if (params.get('error')) {
    const e = new Error(params.get('error_description') || params.get('error'));
    e.code = params.get('error');
    throw e;
  }
  if (!saved || params.get('state') !== saved.state) throw new Error('The sign-in response did not match this browser tab. Please try again.');
  return tokenRequest(saved.tenantId, {
    client_id: saved.clientId,
    grant_type: 'authorization_code',
    code: params.get('code'),
    redirect_uri: redirectUri(),
    code_verifier: saved.verifier,
    scope: (saved.scopes || requestedScopes()).join(' '),
  });
}

function decodeJwt(jwt) {
  try {
    const part = jwt.split('.')[1].replace(/-/g, '+').replace(/_/g, '/');
    return JSON.parse(new TextDecoder().decode(Uint8Array.from(atob(part), (c) => c.charCodeAt(0))));
  } catch { return null; }
}

async function tokenRequest(tenant, fields) {
  const res = await fetch(`${authority(tenant)}/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams(fields),
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) {
    const e = new Error(data.error_description || data.error || `Sign-in failed (${res.status})`);
    e.code = data.error;
    throw e;
  }
  const prev = getTokens();
  const tokens = {
    access_token: data.access_token,
    refresh_token: data.refresh_token || prev?.refresh_token,
    expires_at: Date.now() + (Number(data.expires_in || 3600) - 120) * 1000,
    tenant,
    clientId: fields.client_id,
    account: prev?.account || null,
    scopes: fields.scope.split(' '),
    granted: data.scope ? norm(data.scope) : prev?.granted || null,
  };
  const claims = data.id_token && decodeJwt(data.id_token);
  if (claims) {
    tokens.account = { name: claims.name, username: claims.preferred_username };
    if (claims.preferred_username) store.set('loginHint', claims.preferred_username);
  }
  localStorage.setItem(TOKENS, JSON.stringify(tokens));
  return tokens;
}

export function getTokens() {
  try { return JSON.parse(localStorage.getItem(TOKENS)); } catch { return null; }
}

let refreshing = null;
export async function getAccessToken() {
  const t = getTokens();
  if (!t) throw authRequired();
  if (t.access_token && t.expires_at > Date.now()) return t.access_token;
  if (!t.refresh_token) throw authRequired();
  refreshing ||= tokenRequest(t.tenant || msConfig().tenantId, {
    client_id: t.clientId || msConfig().clientId,
    grant_type: 'refresh_token',
    refresh_token: t.refresh_token,
    scope: (t.scopes || CONFIG.baseScopes).join(' '),
  }).finally(() => { refreshing = null; });
  try {
    return (await refreshing).access_token;
  } catch {
    // Browser-app sessions last about 24 hours; after that Microsoft wants a fresh sign-in.
    localStorage.removeItem(TOKENS);
    throw authRequired();
  }
}

function authRequired() {
  const e = new Error('Your Microsoft sign-in has expired. Please sign in again.');
  e.code = 'auth_required';
  return e;
}

export function signOut() {
  localStorage.removeItem(TOKENS);
}
