// edutraceug-marks-worker
//
// Records assessment marks, computes term reports, generates printable report cards.
// Consumes assessment types, weights, and scoring mode from edutraceug-teachers-worker.
// Owns the term's maxScores map (used only in percentage mode).
//
// Storage:
//   schools/{sid}/marks/{markId}              one cell = one doc, deterministic id
//   schools/{sid}/markSheets/{sheetId}        lock state per (term,class,stream,subject)
//   schools/{sid}/reportHolds/{rosterId}      per-student withhold (fee arrears, etc.)
//   schools/{sid}/reportNotes/{termId}_{rid}  class-teacher comment + conduct
//   schools/{sid}/settings/gradingScale       bands table (percentage mode)
//   schools/{sid}/auditLog/{auto}
//
// Term doc fields this worker writes:
//   maxScores, publishRank
//
// Secrets:
//   ACCOUNT_SERVICE_FIREBASE (falls back to ACCOUNT-SERVICE-FIREBASE)
//   CLOUDINARY_URL
//   CLOUDMERSIVE_API_KEY

import * as XLSX from "xlsx";

// ================================================================
// Constants
// ================================================================

const DEFAULT_PROJECT_ID = "edutrace-ug";

const MARK_STATUSES = ["entered", "absent", "excused", "na"];
const SHEET_STATUSES = ["open", "submitted", "locked"];
const SCORING_MODES = ["percentage", "cbc"];

const DEFAULT_MAX_SCORES_BY_NAME = {
  BOT: 20, CA: 20, AOI: 20, COURSEWORK: 20, MIDTERM: 20, ENDTERM: 100,
};
const FALLBACK_MAX_SCORE = 100;

const DEFAULT_GRADING_SCALE = {
  bands: [
    { min: 80, max: 100, grade: "A", points: 1, descriptor: "Excellent" },
    { min: 70, max: 80, grade: "B", points: 2, descriptor: "Very good" },
    { min: 60, max: 70, grade: "C", points: 3, descriptor: "Good" },
    { min: 50, max: 60, grade: "D", points: 4, descriptor: "Fair" },
    { min: 40, max: 50, grade: "E", points: 5, descriptor: "Weak" },
    { min: 0, max: 40, grade: "F", points: 6, descriptor: "Fail" },
  ],
};

const MAX_FILE_SIZE_MB = 25;
const COMMIT_CHUNK = 400;
const MAX_GRID_ROWS = 500;
const MAX_MARKS_SCAN = 5000;
const MAX_ROSTER_ROWS = 20000;
const MAX_SHEETS = 2000;
const MAX_HOLDS = 2000;
const MAX_TYPES = 100;

const SCAN_URL = "https://api.cloudmersive.com/virus/scan/file";
const SCAN_TIMEOUT_MS = 60_000;

const ACCEPTED_EXTENSIONS = [".xlsx", ".xls", ".xlsm", ".csv"];
const ACCEPTED_MIME = new Set([
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  "application/vnd.ms-excel",
  "application/vnd.ms-excel.sheet.macroenabled.12",
  "text/csv",
  "application/csv",
  "text/plain",
  "application/octet-stream",
  "",
]);

// Header aliases for Excel parsing
const ALIASES = {
  rosterId: ["rosterid", "id", "studentid", "regno", "regnumber", "registration", "lin"],
  firstName: ["firstname", "givenname", "forename", "first", "fname", "christianname"],
  lastName: ["lastname", "surname", "familyname", "secondname", "last", "lname"],
  name: ["name", "studentname", "student", "fullname", "pupil", "pupilname"],
  class: ["class", "classname", "grade", "form", "level"],
  stream: ["stream", "section", "arm", "division"],
};

const ABSENT_TOKENS = new Set(["abs", "absent", "a"]);
const EXCUSED_TOKENS = new Set(["exc", "excused", "e"]);
const NA_TOKENS = new Set(["n/a", "na", "-", "--", "—"]);

// ================================================================
// Errors
// ================================================================

class HttpError extends Error {
  constructor(status, message, extra = {}) {
    super(message);
    this.status = status;
    this.extra = extra;
  }
}

// ================================================================
// HTTP helpers
// ================================================================

const CORS_HEADERS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, POST, PUT, PATCH, DELETE, OPTIONS",
  "Access-Control-Allow-Headers": "Authorization, Content-Type",
  "Access-Control-Max-Age": "86400",
};

function json(body, status = 200) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { "Content-Type": "application/json", "Cache-Control": "no-store", ...CORS_HEADERS },
  });
}

function html(body, status = 200) {
  return new Response(body, {
    status,
    headers: { "Content-Type": "text/html; charset=utf-8", "Cache-Control": "no-store", ...CORS_HEADERS },
  });
}

function binary(body, contentType, extraHeaders = {}) {
  return new Response(body, {
    status: 200,
    headers: { "Content-Type": contentType, "Cache-Control": "no-store", ...CORS_HEADERS, ...extraHeaders },
  });
}

async function safeJson(request) {
  try {
    const body = await request.json();
    if (!body || typeof body !== "object" || Array.isArray(body)) throw new Error("not object");
    return body;
  } catch {
    throw new HttpError(400, "Invalid JSON body");
  }
}

// ================================================================
// Base64 / UTF-8
// ================================================================

function utf8ToBytes(str) { return new TextEncoder().encode(str); }

function bytesToBase64Url(bytes) {
  const arr = new Uint8Array(bytes);
  let bin = "";
  const CHUNK = 0x8000;
  for (let i = 0; i < arr.length; i += CHUNK) {
    bin += String.fromCharCode.apply(null, arr.subarray(i, i + CHUNK));
  }
  return btoa(bin).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}

function b64urlToBytes(s) {
  s = String(s).replace(/-/g, "+").replace(/_/g, "/");
  while (s.length % 4) s += "=";
  const bin = atob(s);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}

function b64urlToString(s) { return new TextDecoder().decode(b64urlToBytes(s)); }

function base64ToBytes(b64) {
  const bin = atob(b64);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}

// ================================================================
// Text helpers
// ================================================================

function str(v) { return v === null || v === undefined ? "" : String(v); }
function enc(v) { return encodeURIComponent(v); }
function toNumberOrNull(v) {
  if (v === null || v === undefined || v === "") return null;
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}
function round1(n) { return Math.round(n * 10) / 10; }
function round2(n) { return Math.round(n * 100) / 100; }

function generateId() {
  const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
  const bytes = crypto.getRandomValues(new Uint8Array(20));
  let s = "";
  for (const b of bytes) s += chars[b % chars.length];
  return s;
}

const EMOJI_RE = /[\u{1F000}-\u{1FAFF}\u{1F1E6}-\u{1F1FF}\u{2600}-\u{27BF}\u{2B00}-\u{2BFF}\u{FE0F}\u{200D}\u{20E3}\u{1F3FB}-\u{1F3FF}]/gu;

function escapeHtml(s) {
  return str(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

function cleanText(input, maxLen, label) {
  if (typeof input !== "string") throw new HttpError(400, `${label} must be text`);
  const t = input.trim().replace(EMOJI_RE, "");
  if (!t) throw new HttpError(400, `${label} is required`);
  if (t.length > maxLen) throw new HttpError(400, `${label} is too long (max ${maxLen})`);
  return t;
}

function cleanOptionalText(input, maxLen, label) {
  if (input === null || input === undefined || input === "") return "";
  return cleanText(input, maxLen, label);
}

function normalizeHeaderCell(v) {
  return String(v ?? "").toLowerCase().replace(/[^a-z0-9]/g, "");
}

// ================================================================
// Deterministic IDs
// ================================================================

async function sha256Hex(s) {
  const buf = await crypto.subtle.digest("SHA-256", utf8ToBytes(s));
  return [...new Uint8Array(buf)].map((b) => b.toString(16).padStart(2, "0")).join("");
}

async function computeMarkId(termId, rosterId, subject, typeId) {
  return (await sha256Hex(`${termId}|${rosterId}|${subject}|${typeId}`)).slice(0, 24);
}
async function computeSheetId(termId, className, stream, subject) {
  return (await sha256Hex(`${termId}|${className}|${stream}|${subject}`)).slice(0, 24);
}

// ================================================================
// Service account / Firebase
// ================================================================

let _saCache = { raw: null, parsed: null };
function getServiceAccount(env) {
  const raw =
    env["ACCOUNT_SERVICE_FIREBASE"] ||
    env["ACCOUNT-SERVICE-FIREBASE"];
  if (!raw) throw new HttpError(500, "Server misconfigured: ACCOUNT_SERVICE_FIREBASE");
  if (_saCache.raw === raw && _saCache.parsed) return _saCache.parsed;
  const parsed = typeof raw === "object" ? raw : JSON.parse(raw);
  _saCache = { raw, parsed };
  return parsed;
}
function getProjectId(env) { return getServiceAccount(env).project_id || DEFAULT_PROJECT_ID; }

let jwksCache = { keys: null, expiry: 0 };
async function getJwks() {
  const now = Date.now();
  if (jwksCache.keys && now < jwksCache.expiry) return jwksCache.keys;
  const res = await fetch("https://www.googleapis.com/service_accounts/v1/jwk/securetoken@system.gserviceaccount.com");
  if (!res.ok) throw new HttpError(502, "JWKS fetch failed");
  const data = await res.json();
  jwksCache = { keys: data.keys, expiry: now + 3600_000 };
  return data.keys;
}

async function verifyIdToken(token, projectId) {
  const parts = String(token || "").split(".");
  if (parts.length !== 3) throw new HttpError(401, "Malformed token");
  const [headerB64, payloadB64, sigB64] = parts;

  let header, payload;
  try {
    header = JSON.parse(b64urlToString(headerB64));
    payload = JSON.parse(b64urlToString(payloadB64));
  } catch {
    throw new HttpError(401, "Malformed token");
  }

  const keys = await getJwks();
  const jwk = keys.find((k) => k.kid === header.kid);
  if (!jwk) throw new HttpError(401, "Unknown signing key");

  const { alg, ...cleanJwk } = jwk;
  const key = await crypto.subtle.importKey(
    "jwk",
    cleanJwk,
    { name: "RSASSA-PKCS1-v1_5", hash: "SHA-256" },
    false,
    ["verify"]
  );
  const ok = await crypto.subtle.verify(
    "RSASSA-PKCS1-v1_5",
    key,
    b64urlToBytes(sigB64),
    utf8ToBytes(`${headerB64}.${payloadB64}`)
  );
  if (!ok) throw new HttpError(401, "Invalid token signature");

  const now = Math.floor(Date.now() / 1000);
  if (payload.exp < now) throw new HttpError(401, "Token expired");
  if (payload.iat > now + 60) throw new HttpError(401, "Token issued in future");
  if (payload.aud !== projectId) throw new HttpError(401, "Invalid audience");
  if (payload.iss !== `https://securetoken.google.com/${projectId}`) {
    throw new HttpError(401, "Invalid issuer");
  }
  if (!payload.sub) throw new HttpError(401, "Missing subject");
  return payload;
}

async function requireUser(request, env, schoolId, allowedRoles) {
  const auth = request.headers.get("Authorization") || "";
  const m = auth.match(/^Bearer\s+(.+)$/i);
  if (!m) throw new HttpError(401, "Missing Bearer token");
  const payload = await verifyIdToken(m[1], getProjectId(env));
  if (!allowedRoles.includes(payload.role)) throw new HttpError(403, "Not allowed for your role");
  if (String(payload.schoolId) !== String(schoolId)) {
    throw new HttpError(403, "Not a member of this school");
  }
  return payload;
}

// ================================================================
// Firestore
// ================================================================

let saTokenCache = new Map();

function pemToArrayBuffer(pem) {
  const b64 = pem
    .replace(/\\n/g, "\n")
    .replace(/-----BEGIN PRIVATE KEY-----/, "")
    .replace(/-----END PRIVATE KEY-----/, "")
    .replace(/[\r\n\s]/g, "");
  const bin = atob(b64);
  const buf = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) buf[i] = bin.charCodeAt(i);
  return buf.buffer;
}

async function importPrivateKey(pem) {
  return crypto.subtle.importKey(
    "pkcs8",
    pemToArrayBuffer(pem),
    { name: "RSASSA-PKCS1-v1_5", hash: "SHA-256" },
    false,
    ["sign"]
  );
}

async function getFirestoreToken(env) {
  const sa = getServiceAccount(env);
  const cached = saTokenCache.get(sa.client_email);
  if (cached && cached.expiry > Date.now() + 60_000) return cached.token;

  const now = Math.floor(Date.now() / 1000);
  const header = { alg: "RS256", typ: "JWT" };
  const claim = {
    iss: sa.client_email,
    scope: "https://www.googleapis.com/auth/datastore",
    aud: "https://oauth2.googleapis.com/token",
    iat: now,
    exp: now + 3600,
  };
  const headerB64 = bytesToBase64Url(utf8ToBytes(JSON.stringify(header)));
  const claimB64 = bytesToBase64Url(utf8ToBytes(JSON.stringify(claim)));
  const toSign = `${headerB64}.${claimB64}`;
  const key = await importPrivateKey(sa.private_key);
  const sigBuf = await crypto.subtle.sign("RSASSA-PKCS1-v1_5", key, utf8ToBytes(toSign));
  const jwt = `${toSign}.${bytesToBase64Url(new Uint8Array(sigBuf))}`;

  const res = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: `grant_type=urn:ietf:params:oauth:grant-type:jwt-bearer&assertion=${jwt}`,
  });
  if (!res.ok) throw new HttpError(502, "Firestore auth failed: " + (await res.text()));
  const data = await res.json();
  saTokenCache.set(sa.client_email, { token: data.access_token, expiry: Date.now() + 3500_000 });
  return data.access_token;
}

function toFsValue(v) {
  if (v === null || v === undefined) return { nullValue: null };
  if (typeof v === "string") return { stringValue: v };
  if (typeof v === "boolean") return { booleanValue: v };
  if (typeof v === "number") return Number.isInteger(v) ? { integerValue: String(v) } : { doubleValue: v };
  if (Array.isArray(v)) return { arrayValue: { values: v.map(toFsValue) } };
  if (typeof v === "object") {
    const fields = {};
    for (const [k, val] of Object.entries(v)) fields[k] = toFsValue(val);
    return { mapValue: { fields } };
  }
  throw new HttpError(500, "Cannot serialize value");
}

function fromFsValue(v) {
  if (!v) return null;
  if ("stringValue" in v) return v.stringValue;
  if ("integerValue" in v) return parseInt(v.integerValue, 10);
  if ("doubleValue" in v) return v.doubleValue;
  if ("booleanValue" in v) return v.booleanValue;
  if ("nullValue" in v) return null;
  if ("timestampValue" in v) return v.timestampValue;
  if ("arrayValue" in v) return (v.arrayValue.values || []).map(fromFsValue);
  if ("mapValue" in v) {
    const out = {};
    for (const [k, val] of Object.entries(v.mapValue.fields || {})) out[k] = fromFsValue(val);
    return out;
  }
  return null;
}

function fromFsDoc(doc) {
  const out = {};
  for (const [k, v] of Object.entries(doc.fields || {})) out[k] = fromFsValue(v);
  return out;
}

function fsBase(env) {
  return `https://firestore.googleapis.com/v1/projects/${getProjectId(env)}/databases/(default)/documents`;
}
function docName(env, segments) {
  return `projects/${getProjectId(env)}/databases/(default)/documents/${segments.map(enc).join("/")}`;
}

async function fsGet(env, docPath) {
  const token = await getFirestoreToken(env);
  const res = await fetch(`${fsBase(env)}/${docPath}`, { headers: { Authorization: `Bearer ${token}` } });
  if (res.status === 404) return null;
  if (!res.ok) throw new HttpError(502, "Firestore read failed");
  return fromFsDoc(await res.json());
}

async function fsPatch(env, docPath, data) {
  const token = await getFirestoreToken(env);
  const keys = Object.keys(data);
  const mask = keys.map((f) => `updateMask.fieldPaths=${enc(f)}`).join("&");
  const body = { fields: {} };
  for (const k of keys) body.fields[k] = toFsValue(data[k]);
  const res = await fetch(`${fsBase(env)}/${docPath}?${mask}`, {
    method: "PATCH",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  if (!res.ok) throw new HttpError(502, "Firestore write failed: " + (await res.text()));
  return fromFsDoc(await res.json());
}

async function fsCreate(env, collectionPath, data, docId) {
  const token = await getFirestoreToken(env);
  const qs = docId ? `?documentId=${enc(docId)}` : "";
  const body = { fields: {} };
  for (const [k, v] of Object.entries(data)) body.fields[k] = toFsValue(v);
  const res = await fetch(`${fsBase(env)}/${collectionPath}${qs}`, {
    method: "POST",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  if (res.status === 409) throw new HttpError(409, "Document already exists");
  if (!res.ok) throw new HttpError(502, "Firestore create failed: " + (await res.text()));
  const created = await res.json();
  return { id: created.name.split("/").pop(), data: fromFsDoc(created) };
}

async function fsDelete(env, docPath) {
  const token = await getFirestoreToken(env);
  const res = await fetch(`${fsBase(env)}/${docPath}`, {
    method: "DELETE",
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok && res.status !== 404) throw new HttpError(502, "Firestore delete failed");
}

async function fsQuery(env, parentPath, structuredQuery) {
  const token = await getFirestoreToken(env);
  const url = parentPath ? `${fsBase(env)}/${parentPath}:runQuery` : `${fsBase(env)}:runQuery`;
  const res = await fetch(url, {
    method: "POST",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: JSON.stringify({ structuredQuery }),
  });
  if (!res.ok) throw new HttpError(502, "Firestore query failed: " + (await res.text()));
  const rows = await res.json();
  const out = [];
  for (const r of Array.isArray(rows) ? rows : []) {
    if (r.document) out.push({ id: r.document.name.split("/").pop(), data: fromFsDoc(r.document) });
  }
  return out;
}

async function firestoreCommit(env, writes) {
  const token = await getFirestoreToken(env);
  const res = await fetch(`${fsBase(env)}:commit`, {
    method: "POST",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: JSON.stringify({ writes }),
  });
  if (!res.ok) throw new HttpError(502, "Firestore commit failed: " + (await res.text()));
  return res.json();
}

async function commitAll(env, writes) {
  for (let i = 0; i < writes.length; i += COMMIT_CHUNK) {
    await firestoreCommit(env, writes.slice(i, i + COMMIT_CHUNK));
  }
}

function fsEq(field, value) {
  let v;
  if (typeof value === "string") v = { stringValue: value };
  else if (typeof value === "boolean") v = { booleanValue: value };
  else if (typeof value === "number") {
    v = Number.isInteger(value) ? { integerValue: String(value) } : { doubleValue: value };
  } else v = { stringValue: String(value) };
  return { fieldFilter: { field: { fieldPath: field }, op: "EQUAL", value: v } };
}

// ================================================================
// State signing (ambiguous rows from Excel)
// ================================================================

async function getStateKey(env) {
  const sa = getServiceAccount(env);
  const material = utf8ToBytes(String(sa.private_key) + "|" + sa.project_id + "|edutrace-marks-v1");
  const digest = await crypto.subtle.digest("SHA-256", material);
  return crypto.subtle.importKey("raw", digest, { name: "HMAC", hash: "SHA-256" }, false, ["sign", "verify"]);
}

async function signState(env, payload) {
  const key = await getStateKey(env);
  const body = bytesToBase64Url(utf8ToBytes(JSON.stringify(payload)));
  const sig = await crypto.subtle.sign("HMAC", key, utf8ToBytes(body));
  return `${body}.${bytesToBase64Url(new Uint8Array(sig))}`;
}

async function verifyState(env, token) {
  if (typeof token !== "string" || !token.includes(".")) throw new HttpError(400, "Malformed state token");
  const [body, sigB64] = token.split(".");
  if (!body || !sigB64) throw new HttpError(400, "Malformed state token");
  const key = await getStateKey(env);
  const valid = await crypto.subtle.verify("HMAC", key, b64urlToBytes(sigB64), utf8ToBytes(body));
  if (!valid) throw new HttpError(400, "Invalid state token signature");
  const payload = JSON.parse(b64urlToString(body));
  if (payload.exp && payload.exp < Math.floor(Date.now() / 1000)) throw new HttpError(400, "State token expired");
  return payload;
}

// ================================================================
// Audit log
// ================================================================

async function logChange(env, schoolId, entry) {
  try {
    await fsCreate(env, `schools/${enc(schoolId)}/auditLog`, { ...entry, at: new Date().toISOString() });
  } catch (err) {
    console.error("Audit log write failed:", err && err.message);
  }
}

// ================================================================
// Assessment types / scoring settings (read from teachers worker)
// ================================================================

async function listAssessmentTypes(env, schoolId, { includeArchived = false } = {}) {
  const rows = await fsQuery(env, `schools/${enc(schoolId)}`, {
    from: [{ collectionId: "assessmentTypes" }],
    limit: MAX_TYPES,
  });
  const types = rows.map((r) => ({
    id: r.id,
    name: r.data.name || "",
    order: typeof r.data.order === "number" ? r.data.order : 100,
    archived: r.data.archived === true,
  }));
  const filtered = includeArchived ? types : types.filter((t) => !t.archived);
  filtered.sort((a, b) => a.order - b.order || a.name.localeCompare(b.name));
  return filtered;
}

async function getScoringSettings(env, schoolId) {
  const doc = await fsGet(env, `schools/${enc(schoolId)}/settings/scoring`);
  const mode = doc && SCORING_MODES.includes(doc.mode) ? doc.mode : "percentage";
  const cbcLevels =
    doc && Array.isArray(doc.cbcLevels) && doc.cbcLevels.length > 0
      ? doc.cbcLevels
      : [
          { level: 1, label: "Beginning", descriptor: "Needs support" },
          { level: 2, label: "Developing", descriptor: "Grasping with guidance" },
          { level: 3, label: "Achieved", descriptor: "Consistently meets standard" },
        ];
  return { mode, cbcLevels };
}

async function getGradingScale(env, schoolId) {
  const doc = await fsGet(env, `schools/${enc(schoolId)}/settings/gradingScale`);
  if (!doc || !Array.isArray(doc.bands) || doc.bands.length === 0) return DEFAULT_GRADING_SCALE;
  return doc;
}

function applyGradingScale(pct, scale) {
  if (pct === null || pct === undefined || !Number.isFinite(pct)) return null;
  const bands = (scale && Array.isArray(scale.bands) ? scale.bands : DEFAULT_GRADING_SCALE.bands)
    .slice()
    .sort((a, b) => b.min - a.min);
  for (const band of bands) if (pct >= band.min) return band;
  return null;
}

// ================================================================
// Term + max scores
// ================================================================

async function getTermOrThrow(env, schoolId, termId) {
  const term = await fsGet(env, `schools/${enc(schoolId)}/terms/${enc(termId)}`);
  if (!term) throw new HttpError(404, "Term not found");
  return term;
}

async function resolveTermId(env, schoolId, requested) {
  if (requested) return requested;
  const rows = await fsQuery(env, `schools/${enc(schoolId)}`, {
    from: [{ collectionId: "terms" }],
    where: fsEq("isCurrent", true),
    limit: 1,
  });
  if (!rows.length) throw new HttpError(404, "No current term set");
  return rows[0].id;
}

function resolveMaxScores(term, types) {
  const custom = term && typeof term.maxScores === "object" && term.maxScores ? term.maxScores : {};
  const out = {};
  for (const t of types) {
    const v = toNumberOrNull(custom[t.id]);
    if (v && v > 0) out[t.id] = v;
    else if (DEFAULT_MAX_SCORES_BY_NAME[t.name.toUpperCase()]) {
      out[t.id] = DEFAULT_MAX_SCORES_BY_NAME[t.name.toUpperCase()];
    } else {
      out[t.id] = FALLBACK_MAX_SCORE;
    }
  }
  return out;
}

function resolveWeights(term, types) {
  const arr = term && Array.isArray(term.weights) ? term.weights : [];
  const out = {};
  for (const t of types) out[t.id] = 0;
  for (const w of arr) {
    if (w && types.some((t) => t.id === w.typeId) && typeof w.weight === "number") {
      out[w.typeId] = w.weight;
    }
  }
  return out;
}

// ================================================================
// Teachers / permissions
// ================================================================

async function loadTeacher(env, schoolId, uid) {
  const doc = await fsGet(env, `schools/${enc(schoolId)}/teachers/${enc(uid)}`);
  return doc || { assignments: [], classTeacherOf: null, permissions: {} };
}

const DEFAULT_PERMISSIONS = { canEnterMarks: true, canUploadHoliday: true, canReadParentReports: true };

async function requirePermission(env, schoolId, uid, key) {
  const t = await loadTeacher(env, schoolId, uid);
  const perms = { ...DEFAULT_PERMISSIONS, ...(t.permissions || {}) };
  if (!perms[key]) throw new HttpError(403, `You do not have the ${key} permission`);
  return t;
}

function assignmentCovers(teacher, className, stream, subject) {
  if (!Array.isArray(teacher.assignments)) return false;
  return teacher.assignments.some(
    (a) => a.class === className && String(a.stream || "") === String(stream || "") && a.subject === subject
  );
}

function classTeacherCovers(teacher, className, stream) {
  const ct = teacher.classTeacherOf;
  if (!ct || typeof ct !== "object") return false;
  if (ct.class !== className) return false;
  if (ct.stream && stream && ct.stream !== stream) return false;
  return true;
}

// ================================================================
// Roster
// ================================================================

async function fetchRosterStudents(env, schoolId, className, stream) {
  const filters = [fsEq("recordType", "student"), fsEq("class", className)];
  if (stream) filters.push(fsEq("stream", stream));
  const rows = await fsQuery(env, `schools/${enc(schoolId)}`, {
    from: [{ collectionId: "roster" }],
    where: filters.length === 1 ? filters[0] : { compositeFilter: { op: "AND", filters } },
    limit: MAX_ROSTER_ROWS,
  });
  rows.sort((a, b) => {
    const ln = str(a.data.lastName).localeCompare(str(b.data.lastName));
    if (ln) return ln;
    return str(a.data.firstName).localeCompare(str(b.data.firstName));
  });
  return rows;
}

function normalizeName(s) {
  return String(s || "").toLowerCase().replace(/[^a-z0-9]/g, "");
}

// ================================================================
// Cloudmersive
// ================================================================

async function scanFileForViruses(env, bytes, filename) {
  if (!env.CLOUDMERSIVE_API_KEY) throw new HttpError(500, "Server misconfigured: CLOUDMERSIVE_API_KEY");

  const form = new FormData();
  form.append("inputFile", new Blob([bytes]), filename || "upload");

  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), SCAN_TIMEOUT_MS);

  let res;
  try {
    res = await fetch(SCAN_URL, {
      method: "POST",
      headers: { Apikey: env.CLOUDMERSIVE_API_KEY, Accept: "application/json" },
      body: form,
      signal: controller.signal,
    });
  } catch (err) {
    throw new HttpError(502, `Virus scan service unreachable: ${err.message}`);
  } finally {
    clearTimeout(timer);
  }
  if (!res.ok) {
    const detail = await res.text().catch(() => "");
    throw new HttpError(502, `Virus scan failed (${res.status}): ${detail.slice(0, 200)}`);
  }
  const data = await res.json().catch(() => ({}));
  const found = Array.isArray(data.FoundViruses) ? data.FoundViruses : [];
  const clean = data.CleanResult !== false && found.length === 0;
  return { clean, foundViruses: found };
}

// ================================================================
// Cloudinary (parsed from CLOUDINARY_URL)
// ================================================================

function parseCloudinaryUrl(env) {
  const url = env.CLOUDINARY_URL;
  if (!url) throw new HttpError(500, "Server misconfigured: CLOUDINARY_URL");
  const m = String(url).match(/^cloudinary:\/\/([^:]+):([^@]+)@(.+)$/);
  if (!m) throw new HttpError(500, "CLOUDINARY_URL is malformed");
  return { apiKey: m[1], apiSecret: m[2], cloudName: m[3] };
}

async function sha1Hex(s) {
  const buf = await crypto.subtle.digest("SHA-1", utf8ToBytes(s));
  return [...new Uint8Array(buf)].map((b) => b.toString(16).padStart(2, "0")).join("");
}

function safeBaseName(name) {
  return (
    String(name || "upload").replace(/\.[^.]+$/, "").replace(/[^a-zA-Z0-9_-]+/g, "_").slice(0, 60) || "upload"
  );
}

async function uploadRawToCloudinary(env, bytes, filename, folder) {
  const { apiKey, apiSecret, cloudName } = parseCloudinaryUrl(env);

  const dot = String(filename || "").lastIndexOf(".");
  const ext = dot > 0 ? filename.slice(dot).toLowerCase() : ".xlsx";
  const publicId = `${safeBaseName(filename)}_${Date.now()}${ext}`;
  const timestamp = Math.floor(Date.now() / 1000);

  const signedParams = { folder, public_id: publicId, timestamp: String(timestamp) };
  const toSign =
    Object.keys(signedParams).sort().map((k) => `${k}=${signedParams[k]}`).join("&") + apiSecret;
  const signature = await sha1Hex(toSign);

  const form = new FormData();
  form.append("file", new Blob([bytes]), filename || "upload");
  form.append("api_key", apiKey);
  form.append("timestamp", signedParams.timestamp);
  form.append("public_id", publicId);
  form.append("folder", folder);
  form.append("signature", signature);

  const res = await fetch(`https://api.cloudinary.com/v1_1/${cloudName}/raw/upload`, {
    method: "POST",
    body: form,
  });
  if (!res.ok) throw new HttpError(502, "Cloudinary upload failed: " + (await res.text()).slice(0, 200));
  const data = await res.json();
  if (!data.secure_url) throw new HttpError(502, "Cloudinary response missing secure_url");
  return { url: data.secure_url, publicId };
}

// ================================================================
// Excel parsing
// ================================================================

function cellStr(values, idx) {
  if (idx === undefined || idx === null || idx < 0) return "";
  const raw = values[idx];
  if (raw === undefined || raw === null) return "";
  return String(raw).replace(/\s+/g, " ").trim();
}

function detectHeaderRow(matrix, activeTypes) {
  const scanLimit = Math.min(matrix.length, 15);
  const typeNamesNormalized = new Map();
  for (const t of activeTypes) typeNamesNormalized.set(normalizeHeaderCell(t.name), t.id);

  let best = null;
  for (let r = 0; r < scanLimit; r += 1) {
    const row = matrix[r] || [];
    const map = {};
    let score = 0;

    for (let c = 0; c < row.length; c += 1) {
      const key = normalizeHeaderCell(row[c]);
      if (!key) continue;

      // Identity columns
      for (const field of ["rosterId", "firstName", "lastName", "name", "class", "stream"]) {
        if (map[field] !== undefined) continue;
        if (ALIASES[field].includes(key)) {
          map[field] = c;
          score += 1;
          break;
        }
      }
      // Assessment type columns
      if (typeNamesNormalized.has(key)) {
        const typeId = typeNamesNormalized.get(key);
        if (!map.types) map.types = {};
        if (map.types[typeId] === undefined) {
          map.types[typeId] = c;
          score += 1;
        }
      }
    }

    const hasIdentity = map.rosterId !== undefined || (map.firstName !== undefined && map.lastName !== undefined) || map.name !== undefined;
    const hasType = map.types && Object.keys(map.types).length > 0;
    if (hasIdentity && hasType && (!best || score > best.score)) {
      best = { headerRowIndex: r, map, score, positional: false };
    }
  }
  if (best) return best;

  // Positional fallback: rosterId or firstName, lastName, then active types in order
  const map = { types: {} };
  let col = 0;
  map.rosterId = 0; col = 1;
  map.lastName = col; col += 1;
  for (const t of activeTypes) {
    map.types[t.id] = col;
    col += 1;
  }
  return { headerRowIndex: -1, map, score: 0, positional: true };
}

function readRowsFromWorkbook(buffer, activeTypes) {
  let wb;
  try {
    wb = XLSX.read(new Uint8Array(buffer), { type: "array", cellDates: false });
  } catch (err) {
    throw new HttpError(400, `Could not read the spreadsheet: ${err.message}. Is the file corrupt?`);
  }
  if (!wb.SheetNames || wb.SheetNames.length === 0) throw new HttpError(400, "Workbook has no sheets");
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", blankrows: false, raw: false });

  if (matrix.length === 0) throw new HttpError(400, "Sheet is empty");

  const columns = detectHeaderRow(matrix, activeTypes);
  const startIndex = columns.headerRowIndex + 1;
  const rows = [];

  for (let i = startIndex; i < matrix.length; i += 1) {
    const raw = matrix[i] || [];
    const blank = !raw.some((v) => String(v ?? "").trim() !== "");
    if (blank) continue;

    const row = {
      excelRow: i + 1,
      rosterId: cellStr(raw, columns.map.rosterId),
      firstName: cellStr(raw, columns.map.firstName),
      lastName: cellStr(raw, columns.map.lastName),
      fullName: cellStr(raw, columns.map.name),
      class: cellStr(raw, columns.map.class),
      stream: cellStr(raw, columns.map.stream),
      cells: {},
    };
    if (columns.map.types) {
      for (const [typeId, colIdx] of Object.entries(columns.map.types)) {
        row.cells[typeId] = cellStr(raw, colIdx);
      }
    }
    rows.push(row);
  }
  if (rows.length === 0) throw new HttpError(400, "No data rows found below the header");
  return { columns, rows };
}

// Interpret a cell value as { status, score, level }
function interpretCell(raw, scoringMode, maxLevel) {
  const v = String(raw ?? "").trim();
  if (v === "") return null;

  const lower = v.toLowerCase();
  if (ABSENT_TOKENS.has(lower)) return { status: "absent", score: null, level: null };
  if (EXCUSED_TOKENS.has(lower)) return { status: "excused", score: null, level: null };
  if (NA_TOKENS.has(v) || NA_TOKENS.has(lower)) return { status: "na", score: null, level: null };

  const num = Number(v);
  if (!Number.isFinite(num)) return { error: `"${v}" is not a number or known marker` };

  if (scoringMode === "cbc") {
    if (!Number.isInteger(num) || num < 1 || num > maxLevel) {
      return { error: `Level must be an integer 1-${maxLevel} (got "${v}")` };
    }
    return { status: "entered", score: null, level: num };
  }
  if (num < 0) return { error: `Score cannot be negative ("${v}")` };
  return { status: "entered", score: num, level: null };
}

// ================================================================
// Mark sheet state
// ================================================================

async function getSheet(env, schoolId, sheetId) {
  return fsGet(env, `schools/${enc(schoolId)}/markSheets/${enc(sheetId)}`);
}

async function assertSheetWritable(env, schoolId, termId, className, stream, subject, user) {
  const sheetId = await computeSheetId(termId, className, stream, subject);
  const sheet = await getSheet(env, schoolId, sheetId);
  if (!sheet) return { sheetId, sheet: null };
  if (sheet.status === "locked") throw new HttpError(409, "This mark sheet is locked. Ask an admin to reopen it.");
  if (sheet.status === "submitted" && user.role !== "schoolAdmin") {
    throw new HttpError(409, "This mark sheet has been submitted. Ask an admin to reopen it before editing.");
  }
  return { sheetId, sheet };
}

// ================================================================
// Grid: manual entry
// ================================================================

async function fetchGridMarks(env, schoolId, termId, className, stream, subject, typeId) {
  const filters = [
    fsEq("termId", termId),
    fsEq("class", className),
    fsEq("subject", subject),
    fsEq("typeId", typeId),
  ];
  if (stream) filters.push(fsEq("stream", stream));
  const rows = await fsQuery(env, `schools/${enc(schoolId)}`, {
    from: [{ collectionId: "marks" }],
    where: { compositeFilter: { op: "AND", filters } },
    limit: MAX_GRID_ROWS,
  });
  const byRoster = new Map();
  for (const r of rows) byRoster.set(r.data.rosterId, { id: r.id, ...r.data });
  return byRoster;
}

async function handleGetGrid(request, env, schoolId, url) {
  const user = await requireUser(request, env, schoolId, ["schoolAdmin", "teacher"]);
  const termId = await resolveTermId(env, schoolId, url.searchParams.get("termId"));
  const className = cleanText(url.searchParams.get("class") || "", 60, "class");
  const stream = url.searchParams.get("stream") ? cleanText(url.searchParams.get("stream"), 60, "stream") : "";
  const subject = cleanText(url.searchParams.get("subject") || "", 120, "subject");
  const typeId = cleanText(url.searchParams.get("typeId") || "", 40, "typeId");

  if (user.role === "teacher") {
    const t = await requirePermission(env, schoolId, String(user.sub), "canEnterMarks");
    if (!assignmentCovers(t, className, stream, subject)) {
      throw new HttpError(403, "You are not assigned to this class/stream/subject");
    }
  }

  const [term, types, scoring] = await Promise.all([
    getTermOrThrow(env, schoolId, termId),
    listAssessmentTypes(env, schoolId),
    getScoringSettings(env, schoolId),
  ]);
  if (!types.some((t) => t.id === typeId)) throw new HttpError(404, `Assessment type "${typeId}" not found or archived`);

  const maxScores = resolveMaxScores(term, types);
  const weights = resolveWeights(term, types);
  const [students, marksByRoster] = await Promise.all([
    fetchRosterStudents(env, schoolId, className, stream),
    fetchGridMarks(env, schoolId, termId, className, stream, subject, typeId),
  ]);

  const sheetId = await computeSheetId(termId, className, stream, subject);
  const sheet = await getSheet(env, schoolId, sheetId);

  let enteredCount = 0;
  const rows = students.map((s) => {
    const m = marksByRoster.get(s.id);
    const status = m ? m.status : "missing";
    if (status !== "missing") enteredCount += 1;
    return {
      rosterId: s.id,
      firstName: s.data.firstName || "",
      lastName: s.data.lastName || "",
      gender: s.data.gender || "",
      markId: m ? m.id : null,
      status,
      score: m && typeof m.score === "number" ? m.score : null,
      level: m && typeof m.level === "number" ? m.level : null,
      maxScore: m && typeof m.maxScore === "number" ? m.maxScore : maxScores[typeId],
      updatedAt: m ? m.updatedAt : null,
      updatedBy: m ? m.updatedBy : null,
    };
  });

  return {
    scoringMode: scoring.mode,
    cbcLevels: scoring.cbcLevels,
    term: { id: termId, name: term.name || "", year: term.year || null },
    class: className,
    stream,
    subject,
    typeId,
    typeName: types.find((t) => t.id === typeId)?.name || typeId,
    maxScore: maxScores[typeId],
    weight: weights[typeId],
    sheet: sheet
      ? { id: sheetId, status: sheet.status, submittedAt: sheet.submittedAt || null, lockedAt: sheet.lockedAt || null }
      : { id: sheetId, status: "open", submittedAt: null, lockedAt: null },
    progress: { entered: enteredCount, total: rows.length },
    rows,
  };
}

async function handlePutGrid(request, env, schoolId) {
  const user = await requireUser(request, env, schoolId, ["schoolAdmin", "teacher"]);
  const body = await safeJson(request);

  const termId = await resolveTermId(env, schoolId, body.termId);
  const className = cleanText(body.class, 60, "class");
  const stream = body.stream ? cleanText(body.stream, 60, "stream") : "";
  const subject = cleanText(body.subject, 120, "subject");
  const typeId = cleanText(body.typeId, 40, "typeId");

  if (user.role === "teacher") {
    const t = await requirePermission(env, schoolId, String(user.sub), "canEnterMarks");
    if (!assignmentCovers(t, className, stream, subject)) throw new HttpError(403, "Not your assignment");
  }

  await assertSheetWritable(env, schoolId, termId, className, stream, subject, user);

  if (!Array.isArray(body.entries) || body.entries.length === 0) throw new HttpError(400, "entries required");
  if (body.entries.length > MAX_GRID_ROWS) throw new HttpError(400, `Max ${MAX_GRID_ROWS} entries`);

  const [term, types, scoring] = await Promise.all([
    getTermOrThrow(env, schoolId, termId),
    listAssessmentTypes(env, schoolId),
    getScoringSettings(env, schoolId),
  ]);
  if (!types.some((t) => t.id === typeId)) throw new HttpError(404, "Assessment type not found");

  const maxScores = resolveMaxScores(term, types);
  const defaultMax = toNumberOrNull(body.maxScore) || maxScores[typeId];
  const maxLevel = scoring.cbcLevels.length;

  const students = await fetchRosterStudents(env, schoolId, className, stream);
  const validRoster = new Set(students.map((s) => s.id));

  const nowIso = new Date().toISOString();
  const uid = String(user.sub);
  const writes = [];
  const auditRows = [];

  for (const e of body.entries) {
    if (!e || typeof e !== "object") throw new HttpError(400, "Each entry must be an object");
    const rosterId = str(e.rosterId).trim();
    if (!rosterId || !validRoster.has(rosterId)) throw new HttpError(400, `rosterId not in class: ${rosterId}`);

    const status = MARK_STATUSES.includes(e.status) ? e.status : null;
    if (!status) throw new HttpError(400, `Invalid status for ${rosterId}`);

    let score = null, level = null, cellMax = null;

    if (scoring.mode === "percentage") {
      cellMax = toNumberOrNull(e.maxScore) || defaultMax;
      if (!(cellMax > 0)) throw new HttpError(400, `Invalid maxScore for ${rosterId}`);
      if (status === "entered") {
        const s = toNumberOrNull(e.score);
        if (s === null || s < 0 || s > cellMax) {
          throw new HttpError(400, `score for ${rosterId} must be between 0 and ${cellMax}`);
        }
        score = s;
      }
    } else {
      if (status === "entered") {
        const l = toNumberOrNull(e.level);
        if (!Number.isInteger(l) || l < 1 || l > maxLevel) {
          throw new HttpError(400, `level for ${rosterId} must be 1-${maxLevel}`);
        }
        level = l;
      }
    }

    const markId = await computeMarkId(termId, rosterId, subject, typeId);
    const data = {
      termId,
      rosterId,
      class: className,
      stream,
      subject,
      typeId,
      scoringMode: scoring.mode,
      status,
      score,
      level,
      maxScore: cellMax,
      enteredBy: uid,
      enteredAt: nowIso,
      updatedBy: uid,
      updatedAt: nowIso,
    };
    writes.push({
      update: {
        name: docName(env, ["schools", schoolId, "marks", markId]),
        fields: Object.fromEntries(Object.entries(data).map(([k, v]) => [k, toFsValue(v)])),
      },
      updateMask: { fieldPaths: Object.keys(data) },
    });
    auditRows.push({ rosterId, markId, status, score, level, maxScore: cellMax });
  }

  await commitAll(env, writes);
  await logChange(env, schoolId, {
    action: "marks.grid.write",
    by: uid, termId, class: className, stream, subject, typeId,
    count: writes.length, entries: auditRows,
  });

  return { status: "ok", updated: writes.length, termId, class: className, stream, subject, typeId };
}

async function handleDeleteMark(request, env, schoolId, markId) {
  const user = await requireUser(request, env, schoolId, ["schoolAdmin", "teacher"]);
  const doc = await fsGet(env, `schools/${enc(schoolId)}/marks/${enc(markId)}`);
  if (!doc) throw new HttpError(404, "Mark not found");

  if (user.role === "teacher") {
    const t = await requirePermission(env, schoolId, String(user.sub), "canEnterMarks");
    if (!assignmentCovers(t, doc.class, doc.stream || "", doc.subject)) throw new HttpError(403, "Not your assignment");
  }
  await assertSheetWritable(env, schoolId, doc.termId, doc.class, doc.stream || "", doc.subject, user);
  await fsDelete(env, `schools/${enc(schoolId)}/marks/${enc(markId)}`);
  await logChange(env, schoolId, { action: "marks.cell.delete", by: String(user.sub), markId, before: doc });
  return { status: "ok" };
}

// ================================================================
// Excel: template download
// ================================================================

async function handleDownloadTemplate(request, env, schoolId, url) {
  await requireUser(request, env, schoolId, ["schoolAdmin", "teacher"]);
  const className = cleanText(url.searchParams.get("class") || "", 60, "class");
  const stream = url.searchParams.get("stream") ? cleanText(url.searchParams.get("stream"), 60, "stream") : "";
  const subject = cleanText(url.searchParams.get("subject") || "", 120, "subject");

  const [types, scoring] = await Promise.all([
    listAssessmentTypes(env, schoolId),
    getScoringSettings(env, schoolId),
  ]);
  if (types.length === 0) throw new HttpError(400, "No active assessment types");

  const students = await fetchRosterStudents(env, schoolId, className, stream);

  const headers = ["rosterId", "firstName", "lastName", ...types.map((t) => t.name)];
  const rows = students.map((s) => [
    s.id,
    s.data.firstName || "",
    s.data.lastName || "",
    ...types.map(() => ""),
  ]);

  const hint = scoring.mode === "cbc"
    ? ["Enter levels (1, 2, 3) or ABS / EXC / N/A"]
    : ["Enter scores or ABS / EXC / N/A"];

  const ws = XLSX.utils.aoa_to_sheet([headers, hint, ...rows]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Marks");
  const buf = XLSX.write(wb, { type: "array", bookType: "xlsx" });

  const filename = `marks_${className}${stream ? "_" + stream : ""}_${subject}.xlsx`.replace(/[^\w.-]+/g, "_");
  return binary(buf, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", {
    "Content-Disposition": `attachment; filename="${filename}"`,
  });
}

// ================================================================
// Excel: upload
// ================================================================

async function handleExcelUpload(request, env, schoolId) {
  const user = await requireUser(request, env, schoolId, ["schoolAdmin", "teacher"]);

  let form;
  try {
    form = await request.formData();
  } catch (err) {
    throw new HttpError(400, `Could not parse multipart body: ${err.message}`);
  }

  const file = form.get("file");
  if (!file || typeof file === "string" || typeof file.arrayBuffer !== "function") {
    throw new HttpError(400, '"file" is required (multipart, field name "file")');
  }
  const className = cleanText(str(form.get("class")), 60, "class");
  const stream = form.get("stream") ? cleanText(str(form.get("stream")), 60, "stream") : "";
  const subject = cleanText(str(form.get("subject")), 120, "subject");
  const requestedTermId = form.get("termId") ? str(form.get("termId")) : null;

  if (user.role === "teacher") {
    const t = await requirePermission(env, schoolId, String(user.sub), "canEnterMarks");
    if (!assignmentCovers(t, className, stream, subject)) throw new HttpError(403, "Not your assignment");
  }

  const termId = await resolveTermId(env, schoolId, requestedTermId);
  await assertSheetWritable(env, schoolId, termId, className, stream, subject, user);

  const ext = (() => {
    const name = file.name || "";
    const dot = name.lastIndexOf(".");
    return dot >= 0 ? name.slice(dot).toLowerCase() : "";
  })();
  if (!ACCEPTED_EXTENSIONS.includes(ext)) {
    throw new HttpError(400, `Unsupported file type "${ext}". Accepted: ${ACCEPTED_EXTENSIONS.join(", ")}`);
  }
  const mime = String(file.type || "").toLowerCase();
  if (!ACCEPTED_MIME.has(mime)) throw new HttpError(400, `Unsupported content type "${file.type}"`);

  const maxBytes = MAX_FILE_SIZE_MB * 1024 * 1024;
  if (file.size > maxBytes) throw new HttpError(400, `File too large (max ${MAX_FILE_SIZE_MB} MB)`);

  const bytes = await file.arrayBuffer();

  // 1. Virus scan (same pipeline as roster/uploads)
  const scan = await scanFileForViruses(env, bytes, file.name || "upload");
  if (!scan.clean) {
    throw new HttpError(400, "File failed the virus scan and was not processed.", {
      foundViruses: scan.foundViruses,
    });
  }

  // 2. Archive raw file to Cloudinary
  let cloudinaryUrl = null;
  try {
    const uploaded = await uploadRawToCloudinary(
      env,
      bytes,
      file.name || "upload.xlsx",
      `edutraceug/marks-uploads/${schoolId}/${className}${stream ? "_" + stream : ""}/${subject}`
    );
    cloudinaryUrl = uploaded.url;
  } catch (err) {
    // Archival is best-effort — the marks data itself is what matters.
    console.error("Cloudinary archive failed:", err && err.message);
  }

  // 3. Parse
  const [types, scoring, term] = await Promise.all([
    listAssessmentTypes(env, schoolId),
    getScoringSettings(env, schoolId),
    getTermOrThrow(env, schoolId, termId),
  ]);
  if (types.length === 0) throw new HttpError(400, "No active assessment types");

  const { rows: parsedRows } = readRowsFromWorkbook(bytes, types);
  const maxScores = resolveMaxScores(term, types);
  const maxLevel = scoring.cbcLevels.length;

  // 4. Match rows to roster
  const students = await fetchRosterStudents(env, schoolId, className, stream);
  const byId = new Map(students.map((s) => [s.id, s]));
  const byName = new Map();
  for (const s of students) {
    const key = `${normalizeName(s.data.firstName)}|${normalizeName(s.data.lastName)}`;
    if (!byName.has(key)) byName.set(key, []);
    byName.get(key).push(s);
  }

  const matched = []; // { row, student }
  const ambiguous = []; // { row, candidates: [rosterId] }
  const notFound = []; // { row, reason }
  const cellErrors = []; // { row, error }

  for (const row of parsedRows) {
    let student = null;

    if (row.rosterId && byId.has(row.rosterId)) {
      student = byId.get(row.rosterId);
    } else {
      let fn = row.firstName, ln = row.lastName;
      if (!fn && !ln && row.fullName) {
        const parts = row.fullName.split(/\s+/).filter(Boolean);
        if (parts.length >= 2) {
          ln = parts[parts.length - 1];
          fn = parts.slice(0, -1).join(" ");
        } else if (parts.length === 1) {
          fn = parts[0];
        }
      }
      if (fn && ln) {
        const key = `${normalizeName(fn)}|${normalizeName(ln)}`;
        const candidates = byName.get(key) || [];
        if (candidates.length === 1) student = candidates[0];
        else if (candidates.length > 1) {
          ambiguous.push({
            excelRow: row.excelRow,
            name: `${row.firstName || fn} ${row.lastName || ln}`.trim(),
            candidates: candidates.map((c) => ({
              rosterId: c.id,
              firstName: c.data.firstName,
              lastName: c.data.lastName,
              class: c.data.class,
              stream: c.data.stream || "",
            })),
          });
          continue;
        }
      }
    }

    if (!student) {
      notFound.push({ excelRow: row.excelRow, rosterId: row.rosterId, name: row.fullName || `${row.firstName} ${row.lastName}`.trim() });
      continue;
    }

    // Validate and interpret each cell
    const cleaned = {};
    let rowHasError = false;
    for (const type of types) {
      const raw = row.cells[type.id];
      if (raw === undefined || raw === "") continue;
      const parsed = interpretCell(raw, scoring.mode, maxLevel);
      if (!parsed) continue;
      if (parsed.error) {
        cellErrors.push({ excelRow: row.excelRow, typeName: type.name, error: parsed.error });
        rowHasError = true;
        continue;
      }
      cleaned[type.id] = parsed;
    }
    if (rowHasError) continue;
    if (Object.keys(cleaned).length === 0) continue;

    matched.push({ student, cells: cleaned });
  }

  // 5. If there are ambiguous rows, return them with a signed state — caller confirms and re-posts
  if (ambiguous.length > 0) {
    const state = await signState(env, {
      termId, className, stream, subject,
      schoolId,
      cloudinaryUrl,
      matched: matched.map((m) => ({ rosterId: m.student.id, cells: m.cells })),
      exp: Math.floor(Date.now() / 1000) + 900,
    });
    return {
      status: "needs_confirmation",
      summary: {
        matched: matched.length,
        ambiguous: ambiguous.length,
        notFound: notFound.length,
        cellErrors: cellErrors.length,
      },
      ambiguous,
      notFound,
      cellErrors,
      state,
    };
  }

  // 6. No ambiguity — write everything
  const writeResult = await writeParsedMarks(env, schoolId, {
    termId, className, stream, subject,
    scoring, types, maxScores,
    matched, cloudinaryUrl,
    uid: String(user.sub),
  });

  return {
    status: "ok",
    ...writeResult,
    summary: {
      matched: matched.length,
      ambiguous: 0,
      notFound: notFound.length,
      cellErrors: cellErrors.length,
    },
    notFound,
    cellErrors,
    cloudinaryUrl,
  };
}

async function writeParsedMarks(env, schoolId, ctx) {
  const { termId, className, stream, subject, scoring, maxScores, matched, cloudinaryUrl, uid } = ctx;
  const nowIso = new Date().toISOString();
  const writes = [];
  const auditRows = [];
  let written = 0;

  for (const entry of matched) {
    const rosterId = entry.student.id;
    for (const [typeId, parsed] of Object.entries(entry.cells)) {
      const markId = await computeMarkId(termId, rosterId, subject, typeId);
      const cellMax = scoring.mode === "percentage" ? maxScores[typeId] : null;

      // Percentage: reject scores above maxScore
      if (scoring.mode === "percentage" && parsed.status === "entered") {
        if (parsed.score > cellMax) {
          throw new HttpError(400, `Score ${parsed.score} exceeds maxScore ${cellMax} (row ${rosterId}, type ${typeId})`);
        }
      }

      const data = {
        termId, rosterId,
        class: className, stream, subject, typeId,
        scoringMode: scoring.mode,
        status: parsed.status,
        score: parsed.score,
        level: parsed.level,
        maxScore: cellMax,
        source: "excel",
        cloudinaryUrl: cloudinaryUrl || null,
        enteredBy: uid,
        enteredAt: nowIso,
        updatedBy: uid,
        updatedAt: nowIso,
      };
      writes.push({
        update: {
          name: docName(env, ["schools", schoolId, "marks", markId]),
          fields: Object.fromEntries(Object.entries(data).map(([k, v]) => [k, toFsValue(v)])),
        },
        updateMask: { fieldPaths: Object.keys(data) },
      });
      auditRows.push({ rosterId, markId, typeId, status: parsed.status, score: parsed.score, level: parsed.level });
      written += 1;
    }
  }

  await commitAll(env, writes);
  await logChange(env, schoolId, {
    action: "marks.excel.write",
    by: uid, termId, class: className, stream, subject,
    count: written, cloudinaryUrl,
  });

  return { written, studentsAffected: matched.length };
}

async function handleExcelConfirm(request, env, schoolId) {
  const user = await requireUser(request, env, schoolId, ["schoolAdmin", "teacher"]);
  const body = await safeJson(request);
  const { state, resolutions } = body || {};
  if (!state) throw new HttpError(400, "state is required");
  if (!resolutions || typeof resolutions !== "object") throw new HttpError(400, "resolutions is required");

  const payload = await verifyState(env, state);
  if (String(payload.schoolId) !== String(schoolId)) throw new HttpError(400, "State token school mismatch");

  const { termId, className, stream, subject, matched, cloudinaryUrl } = payload;

  if (user.role === "teacher") {
    const t = await requirePermission(env, schoolId, String(user.sub), "canEnterMarks");
    if (!assignmentCovers(t, className, stream, subject)) throw new HttpError(403, "Not your assignment");
  }
  await assertSheetWritable(env, schoolId, termId, className, stream, subject, user);

  const [types, scoring, term] = await Promise.all([
    listAssessmentTypes(env, schoolId),
    getScoringSettings(env, schoolId),
    getTermOrThrow(env, schoolId, termId),
  ]);
  const maxScores = resolveMaxScores(term, types);

  // The confirm payload includes the ambiguous rows again so we can resolve them
  // by excelRow → rosterId. We re-fetch the roster to validate.
  const students = await fetchRosterStudents(env, schoolId, className, stream);
  const byId = new Map(students.map((s) => [s.id, s]));

  // The frontend sends { [excelRow]: rosterId }, and the matched rows it kept.
  // If the frontend wants to also pass the ambiguous row cells, it can — we
  // re-interpret here from the initial body.entries if present.
  const extraMatched = [];
  if (Array.isArray(body.ambiguousRows)) {
    for (const row of body.ambiguousRows) {
      const resolvedRosterId = resolutions[String(row.excelRow)];
      if (!resolvedRosterId) continue;
      const student = byId.get(resolvedRosterId);
      if (!student) throw new HttpError(400, `Resolved rosterId ${resolvedRosterId} not in class`);

      const cleaned = {};
      const cells = row.cells || {};
      for (const type of types) {
        const raw = cells[type.id];
        if (raw === undefined || raw === "") continue;
        const parsed = interpretCell(raw, scoring.mode, scoring.cbcLevels.length);
        if (!parsed || parsed.error) continue;
        cleaned[type.id] = parsed;
      }
      if (Object.keys(cleaned).length > 0) {
        extraMatched.push({ student, cells: cleaned });
      }
    }
  }

  // Merge initial matched rows + newly resolved rows
  const allMatched = [...matched.map((m) => ({ student: byId.get(m.rosterId) || { id: m.rosterId, data: {} }, cells: m.cells })), ...extraMatched]
    .filter((m) => m.student);

  if (allMatched.length === 0) throw new HttpError(400, "No rows to write after resolution");

  const result = await writeParsedMarks(env, schoolId, {
    termId, className, stream, subject,
    scoring, types, maxScores,
    matched: allMatched,
    cloudinaryUrl,
    uid: String(user.sub),
  });

  return { status: "ok", ...result };
}

// ================================================================
// Mark sheets (submit / lock / reopen)
// ================================================================

async function handleListSheets(request, env, schoolId, url) {
  await requireUser(request, env, schoolId, ["schoolAdmin", "teacher"]);
  const termId = await resolveTermId(env, schoolId, url.searchParams.get("termId"));
  const rows = await fsQuery(env, `schools/${enc(schoolId)}`, {
    from: [{ collectionId: "markSheets" }],
    where: fsEq("termId", termId),
    limit: MAX_SHEETS,
  });
  return {
    termId,
    sheets: rows.map((r) => ({
      id: r.id,
      class: r.data.class || "",
      stream: r.data.stream || "",
      subject: r.data.subject || "",
      status: r.data.status || "open",
      submittedAt: r.data.submittedAt || null,
      lockedAt: r.data.lockedAt || null,
      updatedAt: r.data.updatedAt || null,
    })),
  };
}

async function setSheetStatus(request, env, schoolId, targetStatus) {
  const allowedRoles = targetStatus === "submitted" ? ["schoolAdmin", "teacher"] : ["schoolAdmin"];
  const user = await requireUser(request, env, schoolId, allowedRoles);
  const body = await safeJson(request);
  const termId = await resolveTermId(env, schoolId, body.termId);
  const className = cleanText(body.class, 60, "class");
  const stream = body.stream ? cleanText(body.stream, 60, "stream") : "";
  const subject = cleanText(body.subject, 120, "subject");

  if (user.role === "teacher") {
    const t = await requirePermission(env, schoolId, String(user.sub), "canEnterMarks");
    if (!assignmentCovers(t, className, stream, subject)) throw new HttpError(403, "Not your assignment");
  }

  const sheetId = await computeSheetId(termId, className, stream, subject);
  const existing = (await getSheet(env, schoolId, sheetId)) || {};
  const nowIso = new Date().toISOString();
  const uid = String(user.sub);

  const update = {
    termId, class: className, stream, subject,
    status: targetStatus,
    updatedAt: nowIso, updatedBy: uid,
  };
  if (targetStatus === "submitted") { update.submittedAt = nowIso; update.submittedBy = uid; }
  if (targetStatus === "locked") { update.lockedAt = nowIso; update.lockedBy = uid; }
  if (targetStatus === "open") { update.reopenedAt = nowIso; update.reopenedBy = uid; }

  await fsPatch(env, `schools/${enc(schoolId)}/markSheets/${enc(sheetId)}`, update);
  await logChange(env, schoolId, {
    action: `sheet.${targetStatus}`,
    by: uid, termId, class: className, stream, subject,
    before: existing.status || null, after: targetStatus,
  });
  return { status: "ok", sheetId, sheetStatus: targetStatus };
}

// ================================================================
// Config: max scores, publish rank
// ================================================================

async function handleGetConfig(request, env, schoolId, url) {
  await requireUser(request, env, schoolId, ["schoolAdmin", "teacher"]);
  const termId = await resolveTermId(env, schoolId, url.searchParams.get("termId"));
  const [term, types, scoring, scale] = await Promise.all([
    getTermOrThrow(env, schoolId, termId),
    listAssessmentTypes(env, schoolId),
    getScoringSettings(env, schoolId),
    getGradingScale(env, schoolId),
  ]);
  return {
    scoringMode: scoring.mode,
    cbcLevels: scoring.cbcLevels,
    term: { id: termId, name: term.name || "", year: term.year || null, isCurrent: term.isCurrent === true },
    assessmentTypes: types,
    maxScores: resolveMaxScores(term, types),
    weights: resolveWeights(term, types),
    publishRank: term.publishRank === true,
    gradingScale: scale,
  };
}

async function handleSetMaxScores(request, env, schoolId) {
  const admin = await requireUser(request, env, schoolId, ["schoolAdmin"]);
  const body = await safeJson(request);
  const termId = await resolveTermId(env, schoolId, body.termId);
  const existing = await getTermOrThrow(env, schoolId, termId);

  if (!body.maxScores || typeof body.maxScores !== "object" || Array.isArray(body.maxScores)) {
    throw new HttpError(400, "maxScores must be an object keyed by typeId");
  }
  const types = await listAssessmentTypes(env, schoolId, { includeArchived: true });
  const typeIds = new Set(types.map((t) => t.id));

  const clean = {};
  for (const [k, v] of Object.entries(body.maxScores)) {
    if (!typeIds.has(k)) throw new HttpError(400, `Unknown assessment type: ${k}`);
    const n = Number(v);
    if (!Number.isFinite(n) || n <= 0 || n > 1000) throw new HttpError(400, `maxScores.${k} must be 1-1000`);
    clean[k] = n;
  }

  const merged = { ...(existing.maxScores || {}), ...clean };
  await fsPatch(env, `schools/${enc(schoolId)}/terms/${enc(termId)}`, {
    maxScores: merged,
    updatedAt: new Date().toISOString(),
    updatedBy: String(admin.sub),
  });
  await logChange(env, schoolId, {
    action: "term.maxScores.set",
    by: String(admin.sub), termId,
    before: existing.maxScores || {}, after: merged,
  });
  return { termId, maxScores: merged };
}

async function handleSetPublishRank(request, env, schoolId) {
  const admin = await requireUser(request, env, schoolId, ["schoolAdmin"]);
  const body = await safeJson(request);
  const termId = await resolveTermId(env, schoolId, body.termId);
  if (typeof body.publishRank !== "boolean") throw new HttpError(400, "publishRank must be boolean");

  await fsPatch(env, `schools/${enc(schoolId)}/terms/${enc(termId)}`, {
    publishRank: body.publishRank,
    updatedAt: new Date().toISOString(),
    updatedBy: String(admin.sub),
  });
  await logChange(env, schoolId, {
    action: "term.publishRank.set",
    by: String(admin.sub), termId, after: body.publishRank,
  });
  return { termId, publishRank: body.publishRank };
}

async function handleSetGradingScale(request, env, schoolId) {
  const admin = await requireUser(request, env, schoolId, ["schoolAdmin"]);
  const body = await safeJson(request);
  if (!Array.isArray(body.bands) || body.bands.length === 0) throw new HttpError(400, "bands must be a non-empty array");

  const clean = body.bands.map((b) => {
    if (!b || typeof b !== "object") throw new HttpError(400, "Invalid band");
    const min = Number(b.min);
    const max = Number(b.max);
    if (!Number.isFinite(min) || !Number.isFinite(max) || min < 0 || max > 100 || min >= max) {
      throw new HttpError(400, "Each band needs 0 <= min < max <= 100");
    }
    return {
      min, max,
      grade: cleanText(b.grade, 10, "band.grade"),
      points: Number.isFinite(Number(b.points)) ? Number(b.points) : null,
      descriptor: cleanOptionalText(b.descriptor, 60, "band.descriptor"),
    };
  });

  await fsPatch(env, `schools/${enc(schoolId)}/settings/gradingScale`, {
    bands: clean,
    updatedAt: new Date().toISOString(),
    updatedBy: String(admin.sub),
  });
  await logChange(env, schoolId, {
    action: "settings.gradingScale.set",
    by: String(admin.sub), after: clean,
  });
  return { bands: clean };
}

// ================================================================
// Report computation
// ================================================================

async function fetchMarksForStudent(env, schoolId, termId, rosterId) {
  return fsQuery(env, `schools/${enc(schoolId)}`, {
    from: [{ collectionId: "marks" }],
    where: {
      compositeFilter: { op: "AND", filters: [fsEq("termId", termId), fsEq("rosterId", rosterId)] },
    },
    limit: MAX_MARKS_SCAN,
  });
}

async function fetchMarksForClass(env, schoolId, termId, className, stream) {
  const filters = [fsEq("termId", termId), fsEq("class", className)];
  if (stream) filters.push(fsEq("stream", stream));
  return fsQuery(env, `schools/${enc(schoolId)}`, {
    from: [{ collectionId: "marks" }],
    where: { compositeFilter: { op: "AND", filters } },
    limit: MAX_MARKS_SCAN,
  });
}

// Percentage-mode aggregation
function computePercentageReport(marks, types, term) {
  const maxScores = resolveMaxScores(term, types);
  const weights = resolveWeights(term, types);

  const bySubject = new Map();
  for (const m of marks) {
    const d = m.data || m;
    if (!d.subject || !d.typeId) continue;
    if (!bySubject.has(d.subject)) bySubject.set(d.subject, new Map());
    bySubject.get(d.subject).set(d.typeId, d);
  }

  const subjects = [];
  let totalPct = 0;
  let subjectsWithData = 0;

  for (const [subject, cellsByType] of bySubject.entries()) {
    const assessments = {};
    let weightedRaw = 0;
    let weightsPresent = 0;

    for (const t of types) {
      const c = cellsByType.get(t.id);
      const w = weights[t.id] ?? 0;
      if (!c) {
        assessments[t.id] = { typeId: t.id, typeName: t.name, status: "missing", score: null, maxScore: maxScores[t.id], weight: w, contribution: 0 };
        continue;
      }
      if (c.status === "entered" && typeof c.score === "number" && c.maxScore > 0) {
        const contribution = (c.score / c.maxScore) * w;
        weightedRaw += contribution;
        weightsPresent += w;
        assessments[t.id] = {
          typeId: t.id, typeName: t.name,
          status: "entered",
          score: c.score, maxScore: c.maxScore, weight: w,
          contribution: round2(contribution),
        };
      } else {
        assessments[t.id] = {
          typeId: t.id, typeName: t.name,
          status: c.status || "missing",
          score: null, maxScore: c.maxScore || maxScores[t.id], weight: w, contribution: 0,
        };
      }
    }

    // Normalize over present weights so partial entry still reads as a meaningful %
    const pct = weightsPresent > 0 ? round1((weightedRaw / weightsPresent) * 100) : null;
    subjects.push({
      name: subject,
      assessments: Object.values(assessments),
      weightedRaw: round2(weightedRaw),
      weightsPresent: round2(weightsPresent),
      pct,
      grade: null,
    });
    if (pct !== null) { totalPct += pct; subjectsWithData += 1; }
  }

  subjects.sort((a, b) => a.name.localeCompare(b.name));
  const overallPct = subjectsWithData > 0 ? round1(totalPct / subjectsWithData) : null;

  return { subjects, overall: { pct: overallPct }, rollup: null };
}

// CBC-mode aggregation — no percentage anywhere
function computeCbcReport(marks, types, cbcLevels) {
  const levels = cbcLevels.slice().sort((a, b) => a.level - b.level);
  const levelByValue = new Map(levels.map((l) => [l.level, l]));
  const maxLevel = levels.length ? levels[levels.length - 1].level : 3;

  const bySubject = new Map();
  for (const m of marks) {
    const d = m.data || m;
    if (!d.subject || !d.typeId) continue;
    if (!bySubject.has(d.subject)) bySubject.set(d.subject, new Map());
    bySubject.get(d.subject).set(d.typeId, d);
  }

  const subjects = [];
  const rollup = { byLevel: new Map() };
  for (const l of levels) rollup.byLevel.set(l.level, 0);

  for (const [subject, cellsByType] of bySubject.entries()) {
    const assessments = [];
    let highest = 0;

    for (const t of types) {
      const c = cellsByType.get(t.id);
      if (!c) {
        assessments.push({ typeId: t.id, typeName: t.name, status: "missing", level: null });
        continue;
      }
      if (c.status === "entered" && typeof c.level === "number") {
        assessments.push({ typeId: t.id, typeName: t.name, status: "entered", level: c.level });
        if (c.level > highest) highest = c.level;
      } else {
        assessments.push({ typeId: t.id, typeName: t.name, status: c.status, level: null });
      }
    }

    const subjectLevel = highest > 0 ? levelByValue.get(highest) || null : null;
    subjects.push({ name: subject, assessments, level: subjectLevel });
    if (subjectLevel) rollup.byLevel.set(subjectLevel.level, (rollup.byLevel.get(subjectLevel.level) || 0) + 1);
  }

  subjects.sort((a, b) => a.name.localeCompare(b.name));

  const rollupRows = levels
    .map((l) => ({ level: l.level, label: l.label, count: rollup.byLevel.get(l.level) || 0 }))
    .filter((r) => r.count > 0);

  const summaryText = rollupRows.length
    ? rollupRows.map((r) => `${r.label}: ${r.count} subject${r.count === 1 ? "" : "s"}`).join(" · ")
    : "No marks recorded yet";

  return { subjects, overall: { pct: null }, rollup: { rows: rollupRows, summaryText } };
}

async function computeRanksForClass(env, schoolId, termId, className, stream, types, term, scoring) {
  const allMarks = await fetchMarksForClass(env, schoolId, termId, className, stream);
  const byStudent = new Map();
  for (const m of allMarks) {
    const rid = m.data.rosterId;
    if (!rid) continue;
    if (!byStudent.has(rid)) byStudent.set(rid, []);
    byStudent.get(rid).push(m.data);
  }

  const ranked = [];
  for (const [rosterId, marks] of byStudent.entries()) {
    if (scoring.mode === "percentage") {
      const { overall } = computePercentageReport(marks, types, term);
      if (overall.pct !== null) ranked.push({ rosterId, key: overall.pct });
    } else {
      // CBC: rank by count of Achieved (highest level)
      const maxLevel = scoring.cbcLevels.length;
      const achieved = marks.filter((m) => m.status === "entered" && m.level === maxLevel).length;
      ranked.push({ rosterId, key: achieved });
    }
  }
  ranked.sort((a, b) => b.key - a.key);

  const positions = new Map();
  let lastKey = null, lastPos = 0;
  ranked.forEach((r, i) => {
    if (r.key !== lastKey) { lastPos = i + 1; lastKey = r.key; }
    positions.set(r.rosterId, lastPos);
  });
  return { positions, totalRanked: ranked.length };
}

async function getReportHold(env, schoolId, rosterId, termId) {
  const hold = await fsGet(env, `schools/${enc(schoolId)}/reportHolds/${enc(rosterId)}`);
  if (!hold) return null;
  if (hold.termId && hold.termId !== termId) return null;
  return hold;
}

async function getReportNote(env, schoolId, termId, rosterId) {
  return fsGet(env, `schools/${enc(schoolId)}/reportNotes/${enc(`${termId}_${rosterId}`)}`);
}

async function buildStudentReport(env, schoolId, rosterId, termId, { includeRank, includeHold }) {
  const [term, student, types, scoring, scale] = await Promise.all([
    getTermOrThrow(env, schoolId, termId),
    fsGet(env, `schools/${enc(schoolId)}/roster/${enc(rosterId)}`),
    listAssessmentTypes(env, schoolId, { includeArchived: true }),
    getScoringSettings(env, schoolId),
    getGradingScale(env, schoolId),
  ]);
  if (!student) throw new HttpError(404, "Student not found in roster");
  if (student.recordType && student.recordType !== "student") throw new HttpError(400, "rosterId is not a student");

  const marksDocs = await fetchMarksForStudent(env, schoolId, termId, rosterId);
  const marks = marksDocs.map((m) => m.data);

  // Include all types that have marks, plus all active types, so empty cells render as blanks.
  const activeTypes = types.filter((t) => !t.archived);
  const typeIdsWithMarks = new Set(marks.map((m) => m.typeId).filter(Boolean));
  const reportTypes = activeTypes.filter((t) => typeIdsWithMarks.has(t.id) || true); // show all active types

  let report;
  if (scoring.mode === "cbc") {
    report = computeCbcReport(marks, reportTypes, scoring.cbcLevels);
  } else {
    report = computePercentageReport(marks, reportTypes, term);
    for (const s of report.subjects) s.grade = applyGradingScale(s.pct, scale);
    report.overall.grade = applyGradingScale(report.overall.pct, scale);
  }

  // List every subject the student takes, even ones with no marks yet.
  const taken = Array.isArray(student.subjects) ? student.subjects : [];
  const have = new Set(report.subjects.map((s) => String(s.name).toLowerCase()));
  for (const name of taken) {
    const key = String(name || "").toLowerCase();
    if (!key || have.has(key)) continue;
    have.add(key);
    const blank = reportTypes.map((t) => scoring.mode === "cbc"
      ? { typeId: t.id, typeName: t.name, status: "missing", level: null }
      : { typeId: t.id, typeName: t.name, status: "missing", score: null, maxScore: null, weight: 0, contribution: 0 });
    report.subjects.push(scoring.mode === "cbc"
      ? { name, assessments: blank, level: null }
      : { name, assessments: blank, weightedRaw: 0, weightsPresent: 0, pct: null, grade: null });
  }
  report.subjects.sort((a, b) => a.name.localeCompare(b.name));

  let rank = null;
  if (includeRank && term.publishRank === true) {
    const { positions, totalRanked } = await computeRanksForClass(
      env, schoolId, termId, student.class, student.stream || "", reportTypes, term, scoring
    );
    rank = { position: positions.get(rosterId) || null, total: totalRanked };
  }

  const hold = includeHold ? await getReportHold(env, schoolId, rosterId, termId) : null;
  const note = await getReportNote(env, schoolId, termId, rosterId);

  return {
    scoringMode: scoring.mode,
    cbcLevels: scoring.cbcLevels,
    term: {
      id: termId,
      name: term.name || "",
      year: term.year || null,
      startDate: term.startDate || "",
      endDate: term.endDate || "",
    },
    student: {
      rosterId,
      firstName: student.firstName || "",
      lastName: student.lastName || "",
      class: student.class || "",
      stream: student.stream || "",
      gender: student.gender || "",
    },
    subjects: report.subjects,
    overall: report.overall,
    rollup: report.rollup,
    rank,
    hold,
    note: note ? { comment: note.comment || "", conduct: note.conduct || "" } : null,
    gradingScale: scoring.mode === "percentage" ? scale : null,
  };
}

// ================================================================
// Report handlers
// ================================================================

async function assertNotHeldForReader(env, schoolId, rosterId, termId, user) {
  if (user.role === "schoolAdmin" || user.role === "teacher") return;
  const hold = await getReportHold(env, schoolId, rosterId, termId);
  if (hold) throw new HttpError(403, `Report withheld: ${hold.reason || "contact the school"}`);
}

async function handleGetMyReport(request, env, schoolId, url) {
  const user = await requireUser(request, env, schoolId, ["student"]);
  const termId = await resolveTermId(env, schoolId, url.searchParams.get("termId"));
  const rosterId = String(user.rosterId || "");
  if (!rosterId) throw new HttpError(400, "Your account is not linked to a roster record");
  await assertNotHeldForReader(env, schoolId, rosterId, termId, user);
  return buildStudentReport(env, schoolId, rosterId, termId, { includeRank: true, includeHold: false });
}

async function handleGetChildReport(request, env, schoolId, url) {
  const user = await requireUser(request, env, schoolId, ["parent"]);
  const termId = await resolveTermId(env, schoolId, url.searchParams.get("termId"));
  const rosterId = String(user.childRosterId || "");
  if (!rosterId) throw new HttpError(400, "Your account is not linked to a child's roster record");

  const consent = await fsGet(env, `schools/${enc(schoolId)}/consent/${enc(String(user.sub))}`);
  if (!consent || consent.consented !== true) throw new HttpError(403, "Parent consent is missing");

  await assertNotHeldForReader(env, schoolId, rosterId, termId, user);
  return buildStudentReport(env, schoolId, rosterId, termId, { includeRank: true, includeHold: false });
}

async function handleGetStudentReport(request, env, schoolId, rosterId, url) {
  const user = await requireUser(request, env, schoolId, ["schoolAdmin", "teacher"]);
  const termId = await resolveTermId(env, schoolId, url.searchParams.get("termId"));
  const student = await fsGet(env, `schools/${enc(schoolId)}/roster/${enc(rosterId)}`);
  if (!student) throw new HttpError(404, "Student not found");

  if (user.role === "teacher") {
    const t = await loadTeacher(env, schoolId, String(user.sub));
    if (!classTeacherCovers(t, student.class, student.stream || "")) {
      throw new HttpError(403, "You are not the class teacher for this student");
    }
  }
  return buildStudentReport(env, schoolId, rosterId, termId, { includeRank: true, includeHold: true });
}

async function handleGetClassReport(request, env, schoolId, url) {
  const user = await requireUser(request, env, schoolId, ["schoolAdmin", "teacher"]);
  const termId = await resolveTermId(env, schoolId, url.searchParams.get("termId"));
  const className = cleanText(url.searchParams.get("class") || "", 60, "class");
  const stream = url.searchParams.get("stream") ? cleanText(url.searchParams.get("stream"), 60, "stream") : "";

  if (user.role === "teacher") {
    const t = await loadTeacher(env, schoolId, String(user.sub));
    if (!classTeacherCovers(t, className, stream)) throw new HttpError(403, "Not your class");
  }

  const [term, students, types, scoring, scale] = await Promise.all([
    getTermOrThrow(env, schoolId, termId),
    fetchRosterStudents(env, schoolId, className, stream),
    listAssessmentTypes(env, schoolId),
    getScoringSettings(env, schoolId),
    getGradingScale(env, schoolId),
  ]);

  const allMarks = await fetchMarksForClass(env, schoolId, termId, className, stream);
  const byStudent = new Map();
  for (const m of allMarks) {
    const rid = m.data.rosterId;
    if (!rid) continue;
    if (!byStudent.has(rid)) byStudent.set(rid, []);
    byStudent.get(rid).push(m.data);
  }

  const rows = students.map((s) => {
    const marks = byStudent.get(s.id) || [];
    let report;
    if (scoring.mode === "cbc") {
      report = computeCbcReport(marks, types, scoring.cbcLevels);
    } else {
      report = computePercentageReport(marks, types, term);
      for (const sub of report.subjects) sub.grade = applyGradingScale(sub.pct, scale);
      report.overall.grade = applyGradingScale(report.overall.pct, scale);
    }
    return {
      rosterId: s.id,
      firstName: s.data.firstName || "",
      lastName: s.data.lastName || "",
      stream: s.data.stream || "",
      subjects: report.subjects,
      overall: report.overall,
      rollup: report.rollup,
    };
  });

  if (term.publishRank === true || user.role === "schoolAdmin") {
    const keyOf = (r) => (scoring.mode === "cbc" ? -(r.rollup?.rows?.[0]?.count || 0) : r.overall.pct);
    const sorted = rows.filter((r) => (scoring.mode === "cbc" ? r.rollup : r.overall.pct !== null))
      .slice()
      .sort((a, b) => (scoring.mode === "cbc" ? (b.rollup?.rows?.[0]?.count || 0) - (a.rollup?.rows?.[0]?.count || 0) : b.overall.pct - a.overall.pct));
    let lastKey = null, lastPos = 0;
    sorted.forEach((r, i) => {
      const k = keyOf(r);
      if (k !== lastKey) { lastPos = i + 1; lastKey = k; }
      r.position = lastPos;
    });
  }

  return {
    scoringMode: scoring.mode,
    cbcLevels: scoring.cbcLevels,
    term: { id: termId, name: term.name || "", year: term.year || null },
    class: className,
    stream,
    publishRank: term.publishRank === true,
    students: rows,
  };
}

// ================================================================
// Report card HTML
// ================================================================

function renderPercentageReportHtml({ report, school }) {
  const schoolName = escapeHtml(school.name || "School");
  const logoUrl = school.logoUrl ? escapeHtml(school.logoUrl) : "";
  const themeColor = /^#[0-9a-fA-F]{6}$/.test(school.themeColor || "") ? school.themeColor : "#3570BC";

  const t = report.term, s = report.student;
  const types = report.subjects[0]?.assessments || [];

  const rowsHtml = report.subjects.map((sub) => {
    const cells = sub.assessments.map((a) => {
      if (a.status === "missing") return `<td class="cell muted">—</td>`;
      if (a.status === "absent") return `<td class="cell muted">ABS</td>`;
      if (a.status === "excused") return `<td class="cell muted">EXC</td>`;
      if (a.status === "na") return `<td class="cell muted">N/A</td>`;
      return `<td class="cell">${escapeHtml(String(a.score))}<span class="max">/${escapeHtml(String(a.maxScore))}</span></td>`;
    }).join("");
    const pct = sub.pct !== null ? sub.pct.toFixed(1) : "—";
    const grade = sub.grade ? escapeHtml(sub.grade.grade) : "—";
    return `<tr><td class="subject">${escapeHtml(sub.name)}</td>${cells}<td class="total">${pct}</td><td class="grade">${grade}</td></tr>`;
  }).join("");

  const overallPct = report.overall.pct !== null ? report.overall.pct.toFixed(1) : "—";
  const overallGrade = report.overall.grade ? escapeHtml(report.overall.grade.grade) : "—";
  const rankHtml = report.rank && report.rank.position ? `<div><label>Position</label><strong>${report.rank.position} of ${report.rank.total}</strong></div>` : "";

  const note = report.note || {};
  const conductHtml = note.conduct ? `<p><b>Conduct:</b> ${escapeHtml(note.conduct)}</p>` : "";

  return `<!DOCTYPE html><html lang="en-UG"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Report Card — ${escapeHtml(s.firstName)} ${escapeHtml(s.lastName)} — ${escapeHtml(t.name)}</title>
<link rel="preconnect" href="https://fonts.googleapis.com"><link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<style>
  :root{--primary:${themeColor};--text-dark:#1A1A1A;--text-muted:#555;--border:#E5E5E5;--bg-light:#F8F9FA}
  *{margin:0;padding:0;box-sizing:border-box;border-radius:0!important}
  body{font-family:'Plus Jakarta Sans',Arial,sans-serif;background:#fff;color:var(--text-dark);padding:24px;line-height:1.5}
  .card{max-width:820px;margin:0 auto;border:2px solid var(--primary)}
  .header{display:flex;align-items:center;gap:20px;padding:22px 26px;border-bottom:3px solid var(--primary)}
  .logo{width:78px;height:78px;object-fit:contain;flex-shrink:0}
  .logo-fallback{width:78px;height:78px;background:var(--primary);color:#fff;display:flex;align-items:center;justify-content:center;font-weight:800;font-size:1.6rem;flex-shrink:0}
  .school-info h1{font-size:1.5rem;font-weight:800;color:var(--primary);line-height:1.2}
  .school-info p{font-size:.8rem;color:var(--text-muted);margin-top:2px}
  .title-bar{background:var(--primary);color:#fff;padding:12px 26px;text-align:center}
  .title-bar h2{font-size:1rem;font-weight:700;text-transform:uppercase;letter-spacing:1.5px}
  .student-info{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;padding:18px 26px;background:var(--bg-light);border-bottom:1px solid var(--border)}
  .student-info div{font-size:.85rem}
  .student-info label{display:block;font-size:.7rem;text-transform:uppercase;letter-spacing:.5px;color:var(--text-muted);font-weight:700;margin-bottom:2px}
  table{width:100%;border-collapse:collapse}
  th,td{padding:9px 8px;border-bottom:1px solid var(--border);font-size:.85rem;text-align:center}
  th{background:var(--primary);color:#fff;font-weight:700;text-transform:uppercase;font-size:.7rem;letter-spacing:.5px;padding:10px 6px}
  th.subject,td.subject{text-align:left;font-weight:600}
  td.cell{font-variant-numeric:tabular-nums}
  td.cell .max{color:var(--text-muted);font-size:.7rem;margin-left:1px}
  td.cell.muted{color:var(--text-muted);font-size:.75rem;font-style:italic}
  td.total{font-weight:800;color:var(--primary);font-variant-numeric:tabular-nums}
  td.grade{font-weight:800}
  .summary{display:grid;grid-template-columns:repeat(3,1fr);border-top:2px solid var(--primary)}
  .summary>div{padding:16px 20px;text-align:center;border-right:1px solid var(--border)}
  .summary>div:last-child{border-right:none}
  .summary label{display:block;font-size:.7rem;color:var(--text-muted);text-transform:uppercase;letter-spacing:.5px;font-weight:700;margin-bottom:4px}
  .summary strong{display:block;font-size:1.35rem;font-weight:800;color:var(--primary)}
  .summary small{display:block;font-size:.75rem;color:var(--text-muted);margin-top:2px}
  .comment-box{padding:20px 26px;border-top:1px solid var(--border)}
  .comment-box h3{font-size:.75rem;text-transform:uppercase;letter-spacing:.5px;color:var(--text-muted);margin-bottom:10px;font-weight:700}
  .comment-box p{font-size:.9rem;min-height:44px}
  .signature{display:flex;justify-content:space-between;gap:30px;padding:26px;border-top:1px solid var(--border)}
  .signature>div{flex:1}
  .signature .line{border-bottom:1px solid #999;height:38px;margin-bottom:6px}
  .signature label{font-size:.7rem;color:var(--text-muted);text-transform:uppercase;letter-spacing:.5px;font-weight:700}
  .footer{background:#111;color:#ccc;padding:12px 26px;text-align:center;font-size:.75rem}
  .no-print{padding:16px;text-align:center}
  .no-print button{padding:10px 22px;background:var(--primary);color:#fff;border:none;font-weight:700;text-transform:uppercase;letter-spacing:.5px;cursor:pointer;font-family:inherit}
  @media print{body{padding:0}.card{border:none}.no-print{display:none}}
  @media(max-width:640px){.student-info{grid-template-columns:repeat(2,1fr)}.summary{grid-template-columns:1fr}th,td{padding:6px 4px;font-size:.75rem}}
</style></head><body>
<div class="no-print"><button onclick="window.print()">Print Report Card</button></div>
<div class="card">
  <div class="header">
    ${logoUrl ? `<img class="logo" src="${logoUrl}" alt="">` : `<div class="logo-fallback">${escapeHtml((schoolName[0] || "S").toUpperCase())}</div>`}
    <div class="school-info"><h1>${schoolName}</h1>
      ${school.location ? `<p>${escapeHtml(school.location)}</p>` : ""}
      ${school.phone ? `<p>Tel: ${escapeHtml(school.phone)}</p>` : ""}
      ${school.email ? `<p>Email: ${escapeHtml(school.email)}</p>` : ""}
    </div>
  </div>
  <div class="title-bar"><h2>Term Report Card — ${escapeHtml(t.name)}${t.year ? ", " + escapeHtml(String(t.year)) : ""}</h2></div>
  <div class="student-info">
    <div><label>Name</label><strong>${escapeHtml(s.firstName)} ${escapeHtml(s.lastName)}</strong></div>
    <div><label>Class</label><strong>${escapeHtml(s.class)}</strong></div>
    <div><label>Stream</label><strong>${escapeHtml(s.stream || "—")}</strong></div>
    <div><label>Gender</label><strong>${escapeHtml(s.gender || "—")}</strong></div>
  </div>
  <table>
    <thead><tr><th class="subject">Subject</th>${types.map((a) => `<th>${escapeHtml(a.typeName)}</th>`).join("")}<th>Total %</th><th>Grade</th></tr></thead>
    <tbody>${rowsHtml || `<tr><td colspan="${types.length + 3}" class="cell muted">No marks recorded for this term.</td></tr>`}</tbody>
  </table>
  <div class="summary">
    <div><label>Overall</label><strong>${overallPct}%</strong></div>
    <div><label>Grade</label><strong>${overallGrade}</strong></div>
    ${rankHtml || `<div><label>Subjects</label><strong>${report.subjects.length}</strong></div>`}
  </div>
  <div class="comment-box"><h3>Class Teacher's Comment</h3>${conductHtml}<p>${escapeHtml(note.comment || "") || "&nbsp;"}</p></div>
  <div class="signature">
    <div><div class="line"></div><label>Class Teacher</label></div>
    <div><div class="line"></div><label>Head Teacher</label></div>
    <div><div class="line"></div><label>Date</label></div>
  </div>
  <div class="footer">© ${new Date().getFullYear()} ${schoolName}. Powered by Edutrace.</div>
</div></body></html>`;
}

function renderCbcReportHtml({ report, school }) {
  const schoolName = escapeHtml(school.name || "School");
  const logoUrl = school.logoUrl ? escapeHtml(school.logoUrl) : "";
  const themeColor = /^#[0-9a-fA-F]{6}$/.test(school.themeColor || "") ? school.themeColor : "#3570BC";

  const t = report.term, s = report.student;
  const types = report.subjects[0]?.assessments || [];
  const levelByValue = new Map(report.cbcLevels.map((l) => [l.level, l]));

  const rowsHtml = report.subjects.map((sub) => {
    const cells = sub.assessments.map((a) => {
      if (a.status === "missing") return `<td class="cell muted">—</td>`;
      if (a.status === "absent") return `<td class="cell muted">ABS</td>`;
      if (a.status === "excused") return `<td class="cell muted">EXC</td>`;
      if (a.status === "na") return `<td class="cell muted">N/A</td>`;
      return `<td class="cell">${a.level}</td>`;
    }).join("");
    const finalLabel = sub.level ? escapeHtml(sub.level.label) : "—";
    return `<tr><td class="subject">${escapeHtml(sub.name)}</td>${cells}<td class="grade">${finalLabel}</td></tr>`;
  }).join("");

  const rollup = report.rollup || { rows: [], summaryText: "" };
  const rollupHtml = rollup.rows.map((r) => `<div><label>${escapeHtml(r.label)}</label><strong>${r.count}</strong><small>subject${r.count === 1 ? "" : "s"}</small></div>`).join("");

  const note = report.note || {};
  const conductHtml = note.conduct ? `<p><b>Conduct:</b> ${escapeHtml(note.conduct)}</p>` : "";
  const legendHtml = report.cbcLevels
    .map((l) => `<span><b>${l.level}</b> = ${escapeHtml(l.label)}</span>`)
    .join(" · ");

  return `<!DOCTYPE html><html lang="en-UG"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Report Card — ${escapeHtml(s.firstName)} ${escapeHtml(s.lastName)} — ${escapeHtml(t.name)}</title>
<link rel="preconnect" href="https://fonts.googleapis.com"><link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<style>
  :root{--primary:${themeColor};--text-dark:#1A1A1A;--text-muted:#555;--border:#E5E5E5;--bg-light:#F8F9FA}
  *{margin:0;padding:0;box-sizing:border-box;border-radius:0!important}
  body{font-family:'Plus Jakarta Sans',Arial,sans-serif;background:#fff;color:var(--text-dark);padding:24px;line-height:1.5}
  .card{max-width:820px;margin:0 auto;border:2px solid var(--primary)}
  .header{display:flex;align-items:center;gap:20px;padding:22px 26px;border-bottom:3px solid var(--primary)}
  .logo{width:78px;height:78px;object-fit:contain;flex-shrink:0}
  .logo-fallback{width:78px;height:78px;background:var(--primary);color:#fff;display:flex;align-items:center;justify-content:center;font-weight:800;font-size:1.6rem;flex-shrink:0}
  .school-info h1{font-size:1.5rem;font-weight:800;color:var(--primary);line-height:1.2}
  .school-info p{font-size:.8rem;color:var(--text-muted);margin-top:2px}
  .title-bar{background:var(--primary);color:#fff;padding:12px 26px;text-align:center}
  .title-bar h2{font-size:1rem;font-weight:700;text-transform:uppercase;letter-spacing:1.5px}
  .student-info{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;padding:18px 26px;background:var(--bg-light);border-bottom:1px solid var(--border)}
  .student-info div{font-size:.85rem}
  .student-info label{display:block;font-size:.7rem;text-transform:uppercase;letter-spacing:.5px;color:var(--text-muted);font-weight:700;margin-bottom:2px}
  table{width:100%;border-collapse:collapse}
  th,td{padding:9px 8px;border-bottom:1px solid var(--border);font-size:.85rem;text-align:center}
  th{background:var(--primary);color:#fff;font-weight:700;text-transform:uppercase;font-size:.7rem;letter-spacing:.5px;padding:10px 6px}
  th.subject,td.subject{text-align:left;font-weight:600}
  td.grade{font-weight:800;color:var(--primary)}
  .summary{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));border-top:2px solid var(--primary)}
  .summary>div{padding:16px 20px;text-align:center;border-right:1px solid var(--border)}
  .summary>div:last-child{border-right:none}
  .summary label{display:block;font-size:.7rem;color:var(--text-muted);text-transform:uppercase;letter-spacing:.5px;font-weight:700;margin-bottom:4px}
  .summary strong{display:block;font-size:1.35rem;font-weight:800;color:var(--primary)}
  .summary small{display:block;font-size:.75rem;color:var(--text-muted);margin-top:2px}
  .legend{padding:10px 26px;background:var(--bg-light);font-size:.75rem;color:var(--text-muted);border-top:1px solid var(--border)}
  .comment-box{padding:20px 26px;border-top:1px solid var(--border)}
  .comment-box h3{font-size:.75rem;text-transform:uppercase;letter-spacing:.5px;color:var(--text-muted);margin-bottom:10px;font-weight:700}
  .comment-box p{font-size:.9rem;min-height:44px}
  .signature{display:flex;justify-content:space-between;gap:30px;padding:26px;border-top:1px solid var(--border)}
  .signature>div{flex:1}
  .signature .line{border-bottom:1px solid #999;height:38px;margin-bottom:6px}
  .signature label{font-size:.7rem;color:var(--text-muted);text-transform:uppercase;letter-spacing:.5px;font-weight:700}
  .footer{background:#111;color:#ccc;padding:12px 26px;text-align:center;font-size:.75rem}
  .no-print{padding:16px;text-align:center}
  .no-print button{padding:10px 22px;background:var(--primary);color:#fff;border:none;font-weight:700;text-transform:uppercase;letter-spacing:.5px;cursor:pointer;font-family:inherit}
  @media print{body{padding:0}.card{border:none}.no-print{display:none}}
  @media(max-width:640px){.student-info{grid-template-columns:repeat(2,1fr)}th,td{padding:6px 4px;font-size:.75rem}}
</style></head><body>
<div class="no-print"><button onclick="window.print()">Print Report Card</button></div>
<div class="card">
  <div class="header">
    ${logoUrl ? `<img class="logo" src="${logoUrl}" alt="">` : `<div class="logo-fallback">${escapeHtml((schoolName[0] || "S").toUpperCase())}</div>`}
    <div class="school-info"><h1>${schoolName}</h1>
      ${school.location ? `<p>${escapeHtml(school.location)}</p>` : ""}
      ${school.phone ? `<p>Tel: ${escapeHtml(school.phone)}</p>` : ""}
      ${school.email ? `<p>Email: ${escapeHtml(school.email)}</p>` : ""}
    </div>
  </div>
  <div class="title-bar"><h2>Term Report Card — ${escapeHtml(t.name)}${t.year ? ", " + escapeHtml(String(t.year)) : ""}</h2></div>
  <div class="student-info">
    <div><label>Name</label><strong>${escapeHtml(s.firstName)} ${escapeHtml(s.lastName)}</strong></div>
    <div><label>Class</label><strong>${escapeHtml(s.class)}</strong></div>
    <div><label>Stream</label><strong>${escapeHtml(s.stream || "—")}</strong></div>
    <div><label>Gender</label><strong>${escapeHtml(s.gender || "—")}</strong></div>
  </div>
  <table>
    <thead><tr><th class="subject">Subject</th>${types.map((a) => `<th>${escapeHtml(a.typeName)}</th>`).join("")}<th>Level</th></tr></thead>
    <tbody>${rowsHtml || `<tr><td colspan="${types.length + 2}" class="cell muted">No marks recorded for this term.</td></tr>`}</tbody>
  </table>
  <div class="legend"><b>Levels:</b> ${legendHtml}</div>
  <div class="summary">${rollupHtml || `<div><label>Subjects</label><strong>0</strong></div>`}</div>
  <div class="comment-box"><h3>Class Teacher's Comment</h3>${conductHtml}<p>${escapeHtml(note.comment || "") || "&nbsp;"}</p></div>
  <div class="signature">
    <div><div class="line"></div><label>Class Teacher</label></div>
    <div><div class="line"></div><label>Head Teacher</label></div>
    <div><div class="line"></div><label>Date</label></div>
  </div>
  <div class="footer">© ${new Date().getFullYear()} ${schoolName}. Powered by Edutrace.</div>
</div></body></html>`;
}

function renderReportCardHtml({ report, school }) {
  return report.scoringMode === "cbc"
    ? renderCbcReportHtml({ report, school })
    : renderPercentageReportHtml({ report, school });
}

async function handleGetReportCard(request, env, schoolId, rosterId, url) {
  const user = await requireUser(request, env, schoolId, ["schoolAdmin", "teacher"]);
  const termId = await resolveTermId(env, schoolId, url.searchParams.get("termId"));
  const student = await fsGet(env, `schools/${enc(schoolId)}/roster/${enc(rosterId)}`);
  if (!student) throw new HttpError(404, "Student not found");
  if (user.role === "teacher") {
    const t = await loadTeacher(env, schoolId, String(user.sub));
    if (!classTeacherCovers(t, student.class, student.stream || "")) throw new HttpError(403, "Not your class");
  }
  const [report, school] = await Promise.all([
    buildStudentReport(env, schoolId, rosterId, termId, { includeRank: true, includeHold: true }),
    fsGet(env, `schools/${enc(schoolId)}`),
  ]);
  return html(renderReportCardHtml({ report, school: school || { name: "School" } }));
}

async function handleGetMyReportCard(request, env, schoolId, url) {
  const user = await requireUser(request, env, schoolId, ["student"]);
  const termId = await resolveTermId(env, schoolId, url.searchParams.get("termId"));
  const rosterId = String(user.rosterId || "");
  if (!rosterId) throw new HttpError(400, "Your account is not linked to a roster record");
  await assertNotHeldForReader(env, schoolId, rosterId, termId, user);
  const [report, school] = await Promise.all([
    buildStudentReport(env, schoolId, rosterId, termId, { includeRank: true, includeHold: false }),
    fsGet(env, `schools/${enc(schoolId)}`),
  ]);
  return html(renderReportCardHtml({ report, school: school || { name: "School" } }));
}

async function handleGetChildReportCard(request, env, schoolId, url) {
  const user = await requireUser(request, env, schoolId, ["parent"]);
  const termId = await resolveTermId(env, schoolId, url.searchParams.get("termId"));
  const rosterId = String(user.childRosterId || "");
  if (!rosterId) throw new HttpError(400, "Your account is not linked to a child's roster record");
  const consent = await fsGet(env, `schools/${enc(schoolId)}/consent/${enc(String(user.sub))}`);
  if (!consent || consent.consented !== true) throw new HttpError(403, "Parent consent is missing");
  await assertNotHeldForReader(env, schoolId, rosterId, termId, user);
  const [report, school] = await Promise.all([
    buildStudentReport(env, schoolId, rosterId, termId, { includeRank: true, includeHold: false }),
    fsGet(env, `schools/${enc(schoolId)}`),
  ]);
  return html(renderReportCardHtml({ report, school: school || { name: "School" } }));
}

// ================================================================
// Notes
// ================================================================

async function handleSetNote(request, env, schoolId) {
  const user = await requireUser(request, env, schoolId, ["schoolAdmin", "teacher"]);
  const body = await safeJson(request);
  const termId = await resolveTermId(env, schoolId, body.termId);
  const rosterId = cleanText(body.rosterId, 40, "rosterId");

  const student = await fsGet(env, `schools/${enc(schoolId)}/roster/${enc(rosterId)}`);
  if (!student) throw new HttpError(404, "Student not found");
  if (user.role === "teacher") {
    const t = await loadTeacher(env, schoolId, String(user.sub));
    if (!classTeacherCovers(t, student.class, student.stream || "")) throw new HttpError(403, "Only the class teacher can write this note");
  }

  const comment = cleanOptionalText(body.comment, 2000, "comment");
  const conduct = cleanOptionalText(body.conduct, 200, "conduct");

  const id = `${termId}_${rosterId}`;
  const data = {
    termId, rosterId,
    class: student.class || "", stream: student.stream || "",
    comment, conduct,
    updatedAt: new Date().toISOString(),
    updatedBy: String(user.sub),
  };
  await fsPatch(env, `schools/${enc(schoolId)}/reportNotes/${enc(id)}`, data);
  await logChange(env, schoolId, { action: "report.note.set", by: String(user.sub), termId, rosterId });
  return { status: "ok", note: data };
}

// ================================================================
// Holds
// ================================================================

async function handleListHolds(request, env, schoolId, url) {
  await requireUser(request, env, schoolId, ["schoolAdmin"]);
  const termId = url.searchParams.get("termId");
  const rows = await fsQuery(env, `schools/${enc(schoolId)}`, {
    from: [{ collectionId: "reportHolds" }],
    ...(termId ? { where: fsEq("termId", termId) } : {}),
    limit: MAX_HOLDS,
  });
  return {
    holds: rows.map((r) => ({
      rosterId: r.id,
      termId: r.data.termId || null,
      reason: r.data.reason || "",
      by: r.data.by || null,
      at: r.data.at || null,
    })),
  };
}

async function handleCreateHold(request, env, schoolId) {
  const admin = await requireUser(request, env, schoolId, ["schoolAdmin"]);
  const body = await safeJson(request);
  const termId = await resolveTermId(env, schoolId, body.termId);
  const rosterId = cleanText(body.rosterId, 40, "rosterId");
  const reason = cleanOptionalText(body.reason, 300, "reason");

  const data = { termId, rosterId, reason, by: String(admin.sub), at: new Date().toISOString() };
  await fsPatch(env, `schools/${enc(schoolId)}/reportHolds/${enc(rosterId)}`, data);
  await logChange(env, schoolId, { action: "report.hold.create", by: String(admin.sub), termId, rosterId, reason });
  return { status: "ok", hold: data };
}

async function handleDeleteHold(request, env, schoolId, rosterId) {
  const admin = await requireUser(request, env, schoolId, ["schoolAdmin"]);
  await fsDelete(env, `schools/${enc(schoolId)}/reportHolds/${enc(rosterId)}`);
  await logChange(env, schoolId, { action: "report.hold.delete", by: String(admin.sub), rosterId });
  return { status: "ok" };
}

// ================================================================
// Audit
// ================================================================

async function handleGetAuditLog(request, env, schoolId, url) {
  await requireUser(request, env, schoolId, ["schoolAdmin"]);
  const requested = parseInt(url.searchParams.get("limit") || "50", 10);
  const limit = Math.min(Math.max(Number.isInteger(requested) ? requested : 50, 1), 100);
  const rows = await fsQuery(env, `schools/${enc(schoolId)}`, {
    from: [{ collectionId: "auditLog" }],
    orderBy: [{ field: { fieldPath: "at" }, direction: "DESCENDING" }],
    limit,
  });
  return { entries: rows.map((r) => ({ id: r.id, ...r.data })) };
}

// ================================================================
// Router
// ================================================================

async function handle(request, env) {
  const url = new URL(request.url);
  const path = url.pathname;
  const method = request.method;

  if (method === "OPTIONS") return new Response(null, { status: 204, headers: CORS_HEADERS });
  if (method === "GET" && path === "/health") {
    return json({ status: "ok", service: "edutraceug-marks-worker" });
  }

  let m;
  const d = decodeURIComponent;
  const S = "([^/]+)";

  // Grid
  if (method === "GET" && (m = path.match(new RegExp(`^/schools/${S}/grid$`)))) {
    return json(await handleGetGrid(request, env, d(m[1]), url));
  }
  if (method === "PUT" && (m = path.match(new RegExp(`^/schools/${S}/grid$`)))) {
    return json(await handlePutGrid(request, env, d(m[1])));
  }
  if (method === "DELETE" && (m = path.match(new RegExp(`^/schools/${S}/marks/${S}$`)))) {
    return json(await handleDeleteMark(request, env, d(m[1]), d(m[2])));
  }

  // Excel
  if (method === "GET" && (m = path.match(new RegExp(`^/schools/${S}/marks/template$`)))) {
    return handleDownloadTemplate(request, env, d(m[1]), url);
  }
  if (method === "POST" && (m = path.match(new RegExp(`^/schools/${S}/marks/upload$`)))) {
    return json(await handleExcelUpload(request, env, d(m[1])));
  }
  if (method === "POST" && (m = path.match(new RegExp(`^/schools/${S}/marks/upload/confirm$`)))) {
    return json(await handleExcelConfirm(request, env, d(m[1])));
  }

  // Sheets
  if (method === "GET" && (m = path.match(new RegExp(`^/schools/${S}/sheets$`)))) {
    return json(await handleListSheets(request, env, d(m[1]), url));
  }
  if (method === "POST" && (m = path.match(new RegExp(`^/schools/${S}/sheets/submit$`)))) {
    return json(await setSheetStatus(request, env, d(m[1]), "submitted"));
  }
  if (method === "POST" && (m = path.match(new RegExp(`^/schools/${S}/sheets/reopen$`)))) {
    return json(await setSheetStatus(request, env, d(m[1]), "open"));
  }
  if (method === "POST" && (m = path.match(new RegExp(`^/schools/${S}/sheets/lock$`)))) {
    return json(await setSheetStatus(request, env, d(m[1]), "locked"));
  }

  // Config
  if (method === "GET" && (m = path.match(new RegExp(`^/schools/${S}/config$`)))) {
    return json(await handleGetConfig(request, env, d(m[1]), url));
  }
  if (method === "PUT" && (m = path.match(new RegExp(`^/schools/${S}/config/max-scores$`)))) {
    return json(await handleSetMaxScores(request, env, d(m[1])));
  }
  if (method === "PUT" && (m = path.match(new RegExp(`^/schools/${S}/config/publish-rank$`)))) {
    return json(await handleSetPublishRank(request, env, d(m[1])));
  }
  if (method === "PUT" && (m = path.match(new RegExp(`^/schools/${S}/config/grading-scale$`)))) {
    return json(await handleSetGradingScale(request, env, d(m[1])));
  }

  // Reports
  if (method === "GET" && (m = path.match(new RegExp(`^/schools/${S}/reports/me$`)))) {
    return json(await handleGetMyReport(request, env, d(m[1]), url));
  }
  if (method === "GET" && (m = path.match(new RegExp(`^/schools/${S}/reports/me/card$`)))) {
    return handleGetMyReportCard(request, env, d(m[1]), url);
  }
  if (method === "GET" && (m = path.match(new RegExp(`^/schools/${S}/reports/child$`)))) {
    return json(await handleGetChildReport(request, env, d(m[1]), url));
  }
  if (method === "GET" && (m = path.match(new RegExp(`^/schools/${S}/reports/child/card$`)))) {
    return handleGetChildReportCard(request, env, d(m[1]), url);
  }
  if (method === "GET" && (m = path.match(new RegExp(`^/schools/${S}/reports/class$`)))) {
    return json(await handleGetClassReport(request, env, d(m[1]), url));
  }
  if (method === "GET" && (m = path.match(new RegExp(`^/schools/${S}/reports/card/${S}$`)))) {
    return handleGetReportCard(request, env, d(m[1]), d(m[2]), url);
  }
  if (method === "GET" && (m = path.match(new RegExp(`^/schools/${S}/reports/student/${S}$`)))) {
    return json(await handleGetStudentReport(request, env, d(m[1]), d(m[2]), url));
  }

  // Notes
  if (method === "PUT" && (m = path.match(new RegExp(`^/schools/${S}/notes$`)))) {
    return json(await handleSetNote(request, env, d(m[1])));
  }

  // Holds
  if (method === "GET" && (m = path.match(new RegExp(`^/schools/${S}/holds$`)))) {
    return json(await handleListHolds(request, env, d(m[1]), url));
  }
  if (method === "POST" && (m = path.match(new RegExp(`^/schools/${S}/holds$`)))) {
    return json(await handleCreateHold(request, env, d(m[1])));
  }
  if (method === "DELETE" && (m = path.match(new RegExp(`^/schools/${S}/holds/${S}$`)))) {
    return json(await handleDeleteHold(request, env, d(m[1]), d(m[2])));
  }

  // Audit
  if (method === "GET" && (m = path.match(new RegExp(`^/schools/${S}/audit-log$`)))) {
    return json(await handleGetAuditLog(request, env, d(m[1]), url));
  }

  return json({ error: "Not found" }, 404);
}

export default {
  async fetch(request, env) {
    try {
      return await handle(request, env);
    } catch (err) {
      if (err instanceof HttpError) return json({ error: err.message, ...err.extra }, err.status);
      console.error("Unhandled error:", err && err.stack ? err.stack : err);
      return json({ error: "Internal server error" }, 500);
    }
  },
};

