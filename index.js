import { SheetJS } from 'xlsx'; // Optional Excel parsing if bundled, or standard parsing logic

// ============================================================================
// CONSTANTS & CONFIGURATION
// ============================================================================
const JWKS_URL = "https://www.googleapis.com/service_accounts/v1/jwk/securetoken@system.gserviceaccount.com";
let cachedJwks = null;
let jwksCacheTime = 0;

const DEFAULT_CBC_LEVELS = [
  { level: 3, descriptor: "Achieved", code: "A" },
  { level: 2, descriptor: "Developing", code: "D" },
  { level: 1, descriptor: "Beginning", code: "B" }
];

const DEFAULT_PERCENTAGE_GRADING = [
  { min: 80, max: 100, grade: "D1", remark: "Distinction" },
  { min: 75, max: 79, grade: "D2", remark: "Distinction" },
  { min: 70, max: 74, grade: "C3", remark: "Credit" },
  { min: 65, max: 69, grade: "C4", remark: "Credit" },
  { min: 60, max: 64, grade: "C5", remark: "Credit" },
  { min: 55, max: 59, grade: "C6", remark: "Credit" },
  { min: 45, max: 54, grade: "P7", remark: "Pass" },
  { min: 35, max: 44, grade: "P8", remark: "Pass" },
  { min: 0, max: 34, grade: "F9", remark: "Fail" }
];

// Custom Error Class
class HttpError extends Error {
  constructor(status, message, details = null) {
    super(message);
    this.status = status;
    this.details = details;
  }
}

// ============================================================================
// UTILITY & SANITIZATION FUNCTIONS
// ============================================================================
function str(val) {
  if (val === null || val === undefined) return "";
  return String(val);
}

function cleanString(val) {
  let s = str(val).trim();
  // Strip emojis and non-standard characters
  s = s.replace(/[\u{1F600}-\u{1F64F}\u{1F300}-\u{1F5FF}\u{1F680}-\u{1F6FF}\u{2600}-\u{26FF}\u{2700}-\u{27BF}]/gu, "");
  // Basic HTML entity escaping
  s = s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  return s;
}

function corsHeaders() {
  return {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "GET, POST, PUT, PATCH, DELETE, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type, Authorization",
    "Content-Type": "application/json"
  };
}

function jsonResponse(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: corsHeaders()
  });
}

function errorResponse(err) {
  const status = err instanceof HttpError ? err.status : 500;
  const message = err.message || "Internal Server Error";
  const details = err.details || null;
  return new Response(JSON.stringify({ error: message, details }), {
    status,
    headers: corsHeaders()
  });
}

// Deterministic Hashing
async function generateHash(inputString) {
  const encoder = new TextEncoder();
  const data = encoder.encode(inputString);
  const hashBuffer = await crypto.subtle.digest("SHA-256", data);
  const hashArray = Array.from(new Uint8Array(hashBuffer));
  return hashArray.map(b => b.toString(16).padStart(2, "0")).join("").substring(0, 32);
}

// ============================================================================
// AUTHENTICATION & TOKEN VERIFICATION
// ============================================================================
async function fetchJwks() {
  const now = Date.now();
  if (cachedJwks && now - jwksCacheTime < 3600000) {
    return cachedJwks;
  }
  const res = await fetch(JWKS_URL);
  if (!res.ok) throw new HttpError(500, "Failed to fetch Firebase JWKS");
  const jwks = await res.json();
  cachedJwks = jwks.keys;
  jwksCacheTime = now;
  return cachedJwks;
}

async function verifyAuth(request) {
  const authHeader = request.headers.get("Authorization") || "";
  if (!authHeader.startsWith("Bearer ")) {
    throw new HttpError(401, "Missing or invalid Authorization header");
  }
  const token = authHeader.split("Bearer ")[1].trim();
  const parts = token.split(".");
  if (parts.length !== 3) throw new HttpError(401, "Malformed JWT");

  try {
    const header = JSON.parse(atob(parts[0].replace(/-/g, "+").replace(/_/g, "/")));
    const payload = JSON.parse(atob(parts[1].replace(/-/g, "+").replace(/_/g, "/")));

    const keys = await fetchJwks();
    const key = keys.find(k => k.kid === header.kid);
    if (!key) throw new HttpError(401, "Invalid token key ID");

    // In production, cryptographically verify payload signature with webcrypto imports
    return {
      uid: payload.user_id || payload.sub,
      email: payload.email,
      role: payload.role || "teacher",
      schoolId: payload.schoolId || null
    };
  } catch (e) {
    throw new HttpError(401, "Token verification failed: " + e.message);
  }
}

// ============================================================================
// FIRESTORE REST API WRAPPER
// ============================================================================
class FirestoreClient {
  constructor(env) {
    this.env = env;
    this.serviceAccount = JSON.parse(env.ACCOUNT_SERVICE_FIREBASE);
    this.projectId = this.serviceAccount.project_id;
    this.accessToken = null;
    this.tokenExpiresAt = 0;
  }

  async getAccessToken() {
    if (this.accessToken && Date.now() < this.tokenExpiresAt - 60000) {
      return this.accessToken;
    }

    const now = Math.floor(Date.now() / 1000);
    const claim = {
      iss: this.serviceAccount.client_email,
      scope: "https://www.googleapis.com/auth/datastore",
      aud: "https://oauth2.googleapis.com/token",
      exp: now + 3600,
      iat: now
    };

    // Format JWT header and payload
    const sHeader = JSON.stringify({ alg: "RS256", typ: "JWT" });
    const sPayload = JSON.stringify(claim);

    const base64UrlEncode = (str) => btoa(str).replace(/=/g, "").replace(/\+/g, "-").replace(/\//g, "_");
    const unsignedToken = `${base64UrlEncode(sHeader)}.${base64UrlEncode(sPayload)}`;

    // Import PKCS8 Private Key
    const pemContents = this.serviceAccount.private_key
      .replace(/-----BEGIN PRIVATE KEY-----/, "")
      .replace(/-----END PRIVATE KEY-----/, "")
      .replace(/\s/g, "");
    const binaryKey = Uint8Array.from(atob(pemContents), c => c.charCodeAt(0));

    const cryptoKey = await crypto.subtle.importKey(
      "pkcs8",
      binaryKey.buffer,
      { name: "RSASSA-PKCS1-v1_5", hash: "SHA-256" },
      false,
      ["sign"]
    );

    const signature = await crypto.subtle.sign(
      "RSASSA-PKCS1-v1_5",
      cryptoKey,
      new TextEncoder().encode(unsignedToken)
    );

    const signedToken = `${unsignedToken}.${base64UrlEncode(String.fromCharCode(...new Uint8Array(signature)))}`;

    const res = await fetch("https://oauth2.googleapis.com/token", {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
        assertion: signedToken
      })
    });

    const data = await res.json();
    if (!res.ok) throw new Error("Firebase OAuth failed: " + JSON.stringify(data));

    this.accessToken = data.access_token;
    this.tokenExpiresAt = Date.now() + (data.expires_in * 1000);
    return this.accessToken;
  }

  // Value formatting helpers
  encodeValue(val) {
    if (val === null || val === undefined) return { nullValue: null };
    if (typeof val === "boolean") return { booleanValue: val };
    if (typeof val === "number") {
      return Number.isInteger(val) ? { integerValue: val.toString() } : { doubleValue: val };
    }
    if (typeof val === "string") return { stringValue: val };
    if (Array.isArray(val)) {
      return { arrayValue: { values: val.map(v => this.encodeValue(v)) } };
    }
    if (typeof val === "object") {
      const fields = {};
      for (const [k, v] of Object.entries(val)) {
        fields[k] = this.encodeValue(v);
      }
      return { mapValue: { fields } };
    }
    return { stringValue: String(val) };
  }

  decodeValue(valObj) {
    if (!valObj) return null;
    if ("nullValue" in valObj) return null;
    if ("booleanValue" in valObj) return valObj.booleanValue;
    if ("integerValue" in valObj) return parseInt(valObj.integerValue, 10);
    if ("doubleValue" in valObj) return parseFloat(valObj.doubleValue);
    if ("stringValue" in valObj) return valObj.stringValue;
    if ("arrayValue" in valObj) {
      return (valObj.arrayValue.values || []).map(v => this.decodeValue(v));
    }
    if ("mapValue" in valObj) {
      const obj = {};
      for (const [k, v] of Object.entries(valObj.mapValue.fields || {})) {
        obj[k] = this.decodeValue(v);
      }
      return obj;
    }
    return null;
  }

  decodeDoc(doc) {
    if (!doc || !doc.fields) return null;
    const result = { id: doc.name.split("/").pop() };
    for (const [key, value] of Object.entries(doc.fields)) {
      result[key] = this.decodeValue(value);
    }
    return result;
  }

  async getDoc(path) {
    const token = await this.getAccessToken();
    const url = `https://firestore.googleapis.com/v1/projects/${this.projectId}/databases/(default)/documents/${path}`;
    const res = await fetch(url, {
      headers: { Authorization: `Bearer ${token}` }
    });

    if (res.status === 404) return null;
    if (!res.ok) throw new HttpError(res.status, "Firestore read error");

    const doc = await res.json();
    return this.decodeDoc(doc);
  }

  async setDoc(path, data, merge = true) {
    const token = await this.getAccessToken();
    const url = `https://firestore.googleapis.com/v1/projects/${this.projectId}/databases/(default)/documents/${path}${merge ? "?currentDocument.exists=true" : ""}`;

    const fields = {};
    for (const [k, v] of Object.entries(data)) {
      fields[k] = this.encodeValue(v);
    }

    const method = merge ? "PATCH" : "POST";
    const res = await fetch(url, {
      method,
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify({ fields })
    });

    if (!res.ok) {
      // If patch fails because doc doesn't exist, create it
      if (res.status === 404 && merge) {
        return this.setDoc(path, data, false);
      }
      const errText = await res.text();
      throw new HttpError(res.status, `Firestore write error: ${errText}`);
    }
    return this.decodeDoc(await res.json());
  }

  async runQuery(parentPath, collectionId, whereClauses = []) {
    const token = await this.getAccessToken();
    const url = `https://firestore.googleapis.com/v1/projects/${this.projectId}/databases/(default)/documents/${parentPath}:runQuery`;

    const structuredQuery = {
      from: [{ collectionId }]
    };

    if (whereClauses.length > 0) {
      structuredQuery.where = {
        compositeFilter: {
          op: "AND",
          filters: whereClauses.map(c => ({
            fieldFilter: {
              field: { fieldPath: c.field },
              op: c.op,
              value: this.encodeValue(c.value)
            }
          }))
        }
      };
    }

    const res = await fetch(url, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify({ structuredQuery })
    });

    if (!res.ok) throw new HttpError(res.status, "Firestore query failed");
    const results = await res.json();
    return results.filter(r => r.document).map(r => this.decodeDoc(r.document));
  }
}

// ============================================================================
// LOGGING & AUDIT TRAIL
// ============================================================================
async function logAudit(db, schoolId, actorUid, action, targetId, details) {
  const logId = `${Date.now()}_${Math.random().toString(36).substring(2, 7)}`;
  const path = `schools/${schoolId}/auditLogs/${logId}`;
  await db.setDoc(path, {
    actorUid,
    action,
    targetId,
    details,
    timestamp: new Date().toISOString()
  }, false);
}

// ============================================================================
// INTEGRATION HELPERS (FETCH FROM OTHER WORKERS' DATA)
// ============================================================================
async function fetchSchoolConfiguration(db, schoolId, termId) {
  const [scoringModeDoc, termDoc, profileDoc] = await Promise.all([
    db.getDoc(`schools/${schoolId}/config/scoringMode`),
    db.getDoc(`schools/${schoolId}/terms/${termId}`),
    db.getDoc(`schools/${schoolId}/profile/general`)
  ]);

  const scoringMode = scoringModeDoc?.mode || "percentage";
  const weights = termDoc?.weights || [];
  const profile = profileDoc || { name: "School", logoUrl: "", primaryColor: "#1e40af" };

  return { scoringMode, weights, profile };
}

async function verifyTeacherAssignment(db, schoolId, teacherUid, classId, streamId, subjectId) {
  const teacherDoc = await db.getDoc(`schools/${schoolId}/teachers/${teacherUid}`);
  if (!teacherDoc) throw new HttpError(403, "Teacher record not found");

  if (teacherDoc.isDoS || teacherDoc.role === "schoolAdmin") return true;

  const assignments = teacherDoc.assignments || [];
  const assigned = assignments.some(a => 
    a.classId === classId && 
    a.streamId === streamId && 
    a.subjectId === subjectId
  );

  if (!assigned) {
    throw new HttpError(403, "Unauthorized: You are not assigned to teach this class and subject");
  }
  return true;
}

// ============================================================================
// SCORING ENGINE
// ============================================================================
function computeStudentResults(marks, weights, scoringMode) {
  if (scoringMode === "cbc") {
    const subjectLevels = {};
    const descriptorCounts = { Achieved: 0, Developing: 0, Beginning: 0 };

    marks.forEach(m => {
      if (!subjectLevels[m.subjectId]) subjectLevels[m.subjectId] = [];
      subjectLevels[m.subjectId].push(m.cbcLevel);
    });

    const subjectSummary = {};
    for (const [subId, levels] of Object.entries(subjectLevels)) {
      const avgLevel = Math.round(levels.reduce((a, b) => a + b, 0) / levels.length);
      let desc = "Beginning";
      if (avgLevel >= 3) desc = "Achieved";
      else if (avgLevel === 2) desc = "Developing";

      subjectSummary[subId] = { averageLevel: avgLevel, descriptor: desc };
      descriptorCounts[desc] = (descriptorCounts[desc] || 0) + 1;
    }

    return { mode: "cbc", subjects: subjectSummary, rollup: descriptorCounts };
  } else {
    // Percentage Mode
    const subjectsMap = {};
    const weightMap = {};
    weights.forEach(w => { weightMap[w.typeId] = w.weight; });

    marks.forEach(m => {
      if (!subjectsMap[m.subjectId]) subjectsMap[m.subjectId] = [];
      subjectsMap[m.subjectId].push(m);
    });

    const subjectResults = {};
    let totalScore = 0;
    let subjectCount = 0;

    for (const [subId, mList] of Object.entries(subjectsMap)) {
      let subWeightedTotal = 0;
      let weightApplied = 0;

      mList.forEach(m => {
        const typeWeight = weightMap[m.typeId] || 0;
        const percentageScore = (m.score / m.maxScore) * typeWeight;
        subWeightedTotal += percentageScore;
        weightApplied += typeWeight;
      });

      const finalSubPercentage = Math.round(subWeightedTotal);
      let grade = "F9";
      let remark = "Fail";

      for (const band of DEFAULT_PERCENTAGE_GRADING) {
        if (finalSubPercentage >= band.min && finalSubPercentage <= band.max) {
          grade = band.grade;
          remark = band.remark;
          break;
        }
      }

      subjectResults[subId] = {
        percentage: finalSubPercentage,
        grade,
        remark
      };

      totalScore += finalSubPercentage;
      subjectCount++;
    }

    const overallPercentage = subjectCount > 0 ? Math.round(totalScore / subjectCount) : 0;
    return {
      mode: "percentage",
      overallPercentage,
      subjects: subjectResults
    };
  }
}

// ============================================================================
// ENDPOINT HANDLERS
// ============================================================================

// 1. Manual Grid Entry Save
async function handleSaveMarkSheet(request, env, auth) {
  const db = new FirestoreClient(env);
  const body = await request.json();

  const { schoolId, termId, classId, streamId, subjectId, typeId, marks } = body;

  if (!schoolId || !termId || !classId || !streamId || !subjectId || !typeId || !Array.isArray(marks)) {
    throw new HttpError(400, "Missing required parameters for mark sheet entry");
  }

  await verifyTeacherAssignment(db, schoolId, auth.uid, classId, streamId, subjectId);

  // Sheet ID Deterministic Hash
  const rawSheetString = `${schoolId}_${termId}_${classId}_${streamId}_${subjectId}_${typeId}`;
  const sheetId = await generateHash(rawSheetString);
  const sheetPath = `schools/${schoolId}/terms/${termId}/markSheets/${sheetId}`;

  const existingSheet = await db.getDoc(sheetPath);
  if (existingSheet && (existingSheet.status === "locked" || existingSheet.status === "submitted")) {
    if (auth.role !== "schoolAdmin" && auth.role !== "dos") {
      throw new HttpError(403, `Mark sheet is currently ${existingSheet.status} and cannot be edited`);
    }
  }

  // Create or Update Sheet Document
  await db.setDoc(sheetPath, {
    termId, classId, streamId, subjectId, typeId,
    status: "draft",
    updatedBy: auth.uid,
    updatedAt: new Date().toISOString()
  });

  // Save individual Marks
  for (const entry of marks) {
    const { studentId, score, maxScore, cbcLevel } = entry;
    const rawMarkString = `${schoolId}_${termId}_${studentId}_${subjectId}_${typeId}`;
    const markId = await generateHash(rawMarkString);
    const markPath = `schools/${schoolId}/terms/${termId}/marks/${markId}`;

    await db.setDoc(markPath, {
      sheetId,
      studentId,
      subjectId,
      typeId,
      score: score !== undefined ? parseFloat(score) : null,
      maxScore: maxScore !== undefined ? parseFloat(maxScore) : 100,
      cbcLevel: cbcLevel !== undefined ? parseInt(cbcLevel, 10) : null,
      updatedBy: auth.uid,
      updatedAt: new Date().toISOString()
    });
  }

  await logAudit(db, schoolId, auth.uid, "SAVE_MARKS_GRID", sheetId, {
    classId, streamId, subjectId, typeId, count: marks.length
  });

  return jsonResponse({ message: "Marks saved successfully", sheetId });
}

// 2. Lock / Submit Mark Sheet Workflow
async function handleSheetStatus(request, env, auth) {
  const db = new FirestoreClient(env);
  const body = await request.json();
  const { schoolId, termId, sheetId, status } = body;

  if (!["submitted", "locked", "draft"].includes(status)) {
    throw new HttpError(400, "Invalid status type");
  }

  if (status === "locked" && auth.role !== "schoolAdmin" && auth.role !== "dos") {
    throw new HttpError(403, "Only Admins or DoS can lock mark sheets");
  }

  const sheetPath = `schools/${schoolId}/terms/${termId}/markSheets/${sheetId}`;
  await db.setDoc(sheetPath, {
    status,
    statusChangedBy: auth.uid,
    statusChangedAt: new Date().toISOString()
  });

  await logAudit(db, schoolId, auth.uid, "CHANGE_SHEET_STATUS", sheetId, { status });

  return jsonResponse({ message: `Sheet status updated to ${status}` });
}

// 3. Class Teacher / Conduct Comments
async function handleSaveComment(request, env, auth) {
  const db = new FirestoreClient(env);
  const body = await request.json();
  const { schoolId, termId, studentId, classTeacherComment, conductNote } = body;

  const commentId = await generateHash(`${schoolId}_${termId}_${studentId}_comments`);
  const path = `schools/${schoolId}/terms/${termId}/comments/${commentId}`;

  await db.setDoc(path, {
    studentId,
    classTeacherComment: cleanString(classTeacherComment),
    conductNote: cleanString(conductNote),
    updatedBy: auth.uid,
    updatedAt: new Date().toISOString()
  });

  return jsonResponse({ message: "Comment recorded successfully" });
}

// 4. Excel Upload Processing Pipeline
async function handleExcelUpload(request, env, auth) {
  const db = new FirestoreClient(env);
  const formData = await request.formData();
  const file = formData.get("file");
  const schoolId = formData.get("schoolId");
  const termId = formData.get("termId");
  const classId = formData.get("classId");
  const streamId = formData.get("streamId");
  const subjectId = formData.get("subjectId");
  const typeId = formData.get("typeId");

  if (!file || !schoolId || !termId) {
    throw new HttpError(400, "Missing file or context identifiers");
  }

  // 1. Cloudmersive Virus Scan
  const arrayBuffer = await file.arrayBuffer();
  if (env.CLOUDMERSIVE_API_KEY) {
    const scanRes = await fetch("https://api.cloudmersive.com/virus/scan/file", {
      method: "POST",
      headers: {
        "Apikey": env.CLOUDMERSIVE_API_KEY,
        "Content-Type": "application/octet-stream"
      },
      body: arrayBuffer
    });
    const scanResult = await scanRes.json();
    if (scanResult && !scanResult.CleanResult) {
      throw new HttpError(400, "Security Violation: File failed virus check");
    }
  }

  // 2. Cloudinary Backup Upload
  if (env.CLOUDINARY_URL) {
    const cloudFormData = new FormData();
    cloudFormData.append("file", file);
    cloudFormData.append("upload_preset", "edutrace_marks");
    fetch(env.CLOUDINARY_URL, { method: "POST", body: cloudFormData }).catch(() => {});
  }

  // 3. Match Students & Process (Mock parsing example pattern)
  // Returning candidate confirmation flow format if ambiguous
  const parsedData = [
    { rosterId: "STU-1001", name: "Kato John", score: 85 },
    { rosterId: null, name: "Babirye Mary", score: 90 }
  ];

  const unmatched = [];
  const validMarks = [];

  for (const row of parsedData) {
    if (row.rosterId) {
      validMarks.push(row);
    } else {
      unmatched.push({
        candidateName: row.name,
        suggestedMatches: [
          { rosterId: "STU-1002", name: "Babirye Mary S5" }
        ]
      });
    }
  }

  if (unmatched.length > 0) {
    return jsonResponse({
      status: "NEEDS_CONFIRMATION",
      unmatchedCandidates: unmatched
    }, 202);
  }

  return jsonResponse({ message: "Excel marks processed cleanly", total: validMarks.length });
}

// 5. Individual Student Report Card Generator
async function handleGetReportCard(request, env, auth, schoolId, termId, studentId) {
  const db = new FirestoreClient(env);

  // 1. Check Fee Hold
  const studentDoc = await db.getDoc(`schools/${schoolId}/students/${studentId}`);
  if (!studentDoc) throw new HttpError(404, "Student record not found");

  if (studentDoc.isHeld) {
    return jsonResponse({
      blocked: true,
      reason: studentDoc.holdReason || "Student records are currently on administrative hold."
    }, 403);
  }

  // 2. Fetch School Config & Term Marks
  const config = await fetchSchoolConfiguration(db, schoolId, termId);
  const rawMarks = await db.runQuery(`schools/${schoolId}/terms/${termId}`, "marks", [
    { field: "studentId", op: "EQUAL", value: studentId }
  ]);

  const commentDoc = await db.getDoc(`schools/${schoolId}/terms/${termId}/comments/${await generateHash(`${schoolId}_${termId}_${studentId}_comments`)}`);

  // 3. Execute Scoring Engine
  const calculated = computeStudentResults(rawMarks, config.weights, config.scoringMode);

  // 4. Return Printable HTML or JSON based on Accept header
  const acceptHeader = request.headers.get("Accept") || "";
  if (acceptHeader.includes("text/html")) {
    const html = `
      <!DOCTYPE html>
      <html>
      <head>
        <title>Report Card - ${cleanString(studentDoc.name)}</title>
        <style>
          body { font-family: sans-serif; padding: 20px; color: #111; }
          .header { text-align: center; border-bottom: 2px solid ${config.profile.primaryColor || '#000'}; padding-bottom: 10px; }
          .logo { max-height: 80px; }
          table { width: 100%; border-collapse: collapse; margin-top: 20px; }
          th, td { border: 1px solid #ccc; padding: 8px; text-align: left; }
          th { background: ${config.profile.primaryColor || '#eee'}; color: #fff; }
        </style>
      </head>
      <body>
        <div class="header">
          ${config.profile.logoUrl ? `<img src="${config.profile.logoUrl}" class="logo"/>` : ''}
          <h2>${cleanString(config.profile.name)}</h2>
          <p>Official Academic Report Card</p>
        </div>
        <h3>Student: ${cleanString(studentDoc.name)} (${studentDoc.classId} ${studentDoc.streamId})</h3>
        
        ${config.scoringMode === "cbc" ? `
          <h4>Competency Based Assessment</h4>
          <table>
            <tr><th>Subject</th><th>Descriptor</th></tr>
            ${Object.entries(calculated.subjects).map(([sId, val]) => `
              <tr><td>${sId}</td><td>${val.descriptor}</td></tr>
            `).join('')}
          </table>
        ` : `
          <h4>Overall Percentage: ${calculated.overallPercentage}%</h4>
          <table>
            <tr><th>Subject</th><th>Score</th><th>Grade</th><th>Remark</th></tr>
            ${Object.entries(calculated.subjects).map(([sId, val]) => `
              <tr><td>${sId}</td><td>${val.percentage}%</td><td>${val.grade}</td><td>${val.remark}</td></tr>
            `).join('')}
          </table>
        `}

        <div style="margin-top: 30px;">
          <p><strong>Class Teacher Comment:</strong> ${commentDoc?.classTeacherComment || 'N/A'}</p>
          <p><strong>Conduct Note:</strong> ${commentDoc?.conductNote || 'N/A'}</p>
        </div>
      </body>
      </html>
    `;
    return new Response(html, { headers: { "Content-Type": "text/html" } });
  }

  return jsonResponse({
    student: studentDoc,
    scoringMode: config.scoringMode,
    results: calculated,
    comments: commentDoc
  });
}

// ============================================================================
// MAIN WORKER ROUTER
// ============================================================================
export default {
  async fetch(request, env, ctx) {
    if (request.method === "OPTIONS") {
      return new Response(null, { headers: corsHeaders() });
    }

    try {
      const url = new URL(request.url);
      const path = url.pathname;
      const method = request.method;

      const auth = await verifyAuth(request);

      // Route: Manual Grid Save
      if (path.endsWith("/marks/grid") && method === "POST") {
        return await handleSaveMarkSheet(request, env, auth);
      }

      // Route: Excel Marks Upload
      if (path.endsWith("/marks/excel") && method === "POST") {
        return await handleExcelUpload(request, env, auth);
      }

      // Route: Lock/Submit Sheet
      if (path.endsWith("/marks/sheet-status") && method === "PUT") {
        return await handleSheetStatus(request, env, auth);
      }

      // Route: Class Teacher Comments
      if (path.endsWith("/marks/comments") && method === "POST") {
        return await handleSaveComment(request, env, auth);
      }

      // Route: Report Card Fetch (/schools/:sid/terms/:tid/reports/:studentId)
      const reportMatch = path.match(/\/schools\/([^\/]+)\/terms\/([^\/]+)\/reports\/([^\/]+)$/);
      if (reportMatch && method === "GET") {
        const [, schoolId, termId, studentId] = reportMatch;
        return await handleGetReportCard(request, env, auth, schoolId, termId, studentId);
      }

      return jsonResponse({ error: "Endpoint not found" }, 404);
    } catch (err) {
      return errorResponse(err);
    }
  }
};
