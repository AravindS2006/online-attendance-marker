import express from 'express';
import cors from 'cors';
import { randomUUID } from 'crypto';
import path from 'path';
import { fileURLToPath } from 'url';
import fs from 'fs';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const cachePath = process.env.VERCEL 
  ? '/tmp/roster_cache.json'
  : path.join(__dirname, 'roster_cache.json');

const app = express();
app.use(cors());
app.use(express.json());

const PORT = process.env.PORT || 5000;

// ==========================================
// 1. MOCK DATABASES (Sairam College ERP) - Switch to 'let' to allow live Google Sheet syncs
// ==========================================

// Default Faculty Directory (Sairam ERP Administrator - Loads dynamically from environment variables for Vercel)
let FACULTY_DIRECTORY = [
  { 
    username: process.env.ADMIN_USERNAME || "admin", 
    password: process.env.ADMIN_PASSWORD || "adminpassword", 
    name: "ERP Administrator", 
    dept: "ALL" 
  }
];

// Default Classrooms Directory (Initial defaults - loads dynamically from Roster cache / sync)
let CLASSROOMS_DIRECTORY = [
  { classCode: "G3103", latitude: 12.9602, longitude: 80.0570, radius: 30 }
];

// Student Directory (Loads from disk cache or syncs live from Google Sheet)
let COLLEGE_STUDENTS_DIRECTORY = [];

// Load Roster from Disk cache on boot if it exists (protects against node watch auto-restarts)
let lastCacheLoadTime = 0;

function reloadDirectoryCacheIfNeeded() {
  try {
    if (fs.existsSync(cachePath)) {
      const stats = fs.statSync(cachePath);
      if (stats.mtimeMs > lastCacheLoadTime) {
        const cacheData = JSON.parse(fs.readFileSync(cachePath, 'utf8'));
        if (Array.isArray(cacheData.faculty) && Array.isArray(cacheData.students)) {
          FACULTY_DIRECTORY = cacheData.faculty;
          COLLEGE_STUDENTS_DIRECTORY = cacheData.students;
          if (Array.isArray(cacheData.classrooms)) {
            CLASSROOMS_DIRECTORY = cacheData.classrooms;
          }
          lastCacheLoadTime = stats.mtimeMs;
          console.log(`[DB Cache Auto-Reload] Synced directory updated from disk (${FACULTY_DIRECTORY.length} faculty, ${COLLEGE_STUDENTS_DIRECTORY.length} students).`);
        }
      }
    }
  } catch (err) {
    console.error("[DB Cache Auto-Reload Error]", err.message);
  }
}

if (fs.existsSync(cachePath)) {
  try {
    const stats = fs.statSync(cachePath);
    const cacheData = JSON.parse(fs.readFileSync(cachePath, 'utf8'));
    if (Array.isArray(cacheData.faculty) && cacheData.faculty.length > 0 && Array.isArray(cacheData.students) && cacheData.students.length > 0) {
      FACULTY_DIRECTORY = cacheData.faculty;
      COLLEGE_STUDENTS_DIRECTORY = cacheData.students;
      if (Array.isArray(cacheData.classrooms)) {
        CLASSROOMS_DIRECTORY = cacheData.classrooms;
      }
      lastCacheLoadTime = stats.mtimeMs;
      console.log(`[DB Cache] Restored synchronized roster from disk cache (${COLLEGE_STUDENTS_DIRECTORY.length} students, ${FACULTY_DIRECTORY.length} faculty, ${CLASSROOMS_DIRECTORY.length} classrooms).`);
    }
  } catch (err) {
    console.error("[DB Cache Error] Failed to read roster disk cache:", err.message);
  }
}

// In-memory sessions storage
const sessions = {};

// ==========================================
// 2. HELPER FUNCTIONS
// ==========================================

const cleanRegNo = (regNo) => regNo ? String(regNo).replace(/[^a-zA-Z0-9]/g, '').toLowerCase().trim() : '';

// Haversine formula to compute distance in meters between two lat/lng coordinates
function getHaversineDistance(lat1, lon1, lat2, lon2) {
  const R = 6371e3; // Earth radius in meters
  const phi1 = (lat1 * Math.PI) / 180;
  const phi2 = (lat2 * Math.PI) / 180;
  const deltaPhi = ((lat2 - lat1) * Math.PI) / 180;
  const deltaLambda = ((lon2 - lon1) * Math.PI) / 180;

  const a =
    Math.sin(deltaPhi / 2) * Math.sin(deltaPhi / 2) +
    Math.cos(phi1) * Math.cos(phi2) * Math.sin(deltaLambda / 2) * Math.sin(deltaLambda / 2);
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));

  return R * c; // Distance in meters
}

// Generate a random 6-digit numeric OTP for dynamic marking verification
function generateOtp() {
  return Math.floor(100000 + Math.random() * 900000).toString();
}

// Generate a unique 5-digit numeric Class Access Code
function generateUniqueClassCode() {
  let code;
  do {
    code = Math.floor(10000 + Math.random() * 90000).toString();
  } while (sessions[code]);
  return code;
}

// Rotate session OTP and token
function rotateSessionOtp(classCode) {
  const session = sessions[classCode];
  if (!session) return;

  // Push current values to history
  if (session.currentOtp) {
    session.previousOtps.push(session.currentOtp);
    if (session.previousOtps.length > 2) session.previousOtps.shift();
  }
  if (session.currentToken) {
    session.previousTokens.push(session.currentToken);
    if (session.previousTokens.length > 2) session.previousTokens.shift();
  }

  // Generate new OTP and token
  session.currentOtp = generateOtp();
  session.currentToken = randomUUID();
  session.otpUpdatedAt = Date.now();
}

// Start rotation interval for a session
function startSessionInterval(classCode) {
  const intervalId = setInterval(() => {
    if (!sessions[classCode]) {
      clearInterval(intervalId);
      return;
    }
    rotateSessionOtp(classCode);
  }, 20000); // Rotates every 20 seconds

  sessions[classCode].intervalId = intervalId;
}

// ==========================================
// 3. API ENDPOINTS
// ==========================================

// Teacher Authentication Login
app.post('/api/faculty/login', (req, res) => {
  reloadDirectoryCacheIfNeeded();
  const { username, password } = req.body;
  
  if (!username || !password) {
    return res.status(400).json({ error: "Faculty ID and Password are required" });
  }

  const cleanUsername = String(username).replace(/[\r\n\t]/g, '').toLowerCase().trim();
  const cleanPassword = String(password).replace(/[\r\n\t]/g, '').trim();

  const faculty = FACULTY_DIRECTORY.find(
    f => f && f.username && 
         String(f.username).replace(/[\r\n\t]/g, '').toLowerCase().trim() === cleanUsername && 
         String(f.password).replace(/[\r\n\t]/g, '').trim() === cleanPassword
  );

  if (!faculty) {
    return res.status(401).json({ error: "Invalid Sairam Faculty ID or password" });
  }

  res.json({ 
    success: true, 
    name: faculty.name, 
    department: faculty.dept 
  });
});

// Teacher Password Change
app.post('/api/faculty/change-password', (req, res) => {
  const { username, currentPassword, newPassword } = req.body;
  if (!username || !currentPassword || !newPassword) {
    return res.status(400).json({ error: "Username, current password, and new password are required" });
  }

  reloadDirectoryCacheIfNeeded();
  const faculty = FACULTY_DIRECTORY.find(
    f => f && f.username && String(f.username).replace(/[\r\n\t]/g, '').toLowerCase().trim() === String(username).replace(/[\r\n\t]/g, '').toLowerCase().trim()
  );

  if (!faculty) {
    return res.status(404).json({ error: "Faculty account not found" });
  }

  if (String(faculty.password).replace(/[\r\n\t]/g, '').trim() !== String(currentPassword).replace(/[\r\n\t]/g, '').trim()) {
    return res.status(401).json({ error: "Incorrect current password" });
  }

  faculty.password = newPassword;
  console.log(`[Faculty DB] Password updated for teacher: ${faculty.username}`);
  res.json({ success: true, message: "Password updated successfully!" });
});

// Search students in the database (Crash-proof and case-insensitive)
app.get('/api/students/search', (req, res) => {
  const { q } = req.query;
  if (!q) return res.json([]);
  
  const query = q.toString().toLowerCase().trim();
  const results = COLLEGE_STUDENTS_DIRECTORY.filter(s => {
    const reg = s && s.regNo ? cleanRegNo(s.regNo) : "";
    const name = s && s.name ? String(s.name).toLowerCase().trim() : "";
    return reg.includes(query) || name.includes(query);
  });
  
  res.json(results);
});

// Lookup class session by code (Student manual entry check)
app.get('/api/sessions/lookup/:code', (req, res) => {
  const codeNormalized = req.params.code.toString().toUpperCase().trim();
  const session = sessions[codeNormalized];
  if (!session) {
    return res.json({ found: false });
  }
  res.json({
    found: true,
    classroomCode: session.classroomCode,
    department: session.department,
    section: session.section,
    year: session.year,
    teacherName: session.teacherName || "ERP Administrator",
    geofence: session.geofence
  });
});

// Create a new attendance session (Teacher)
app.post('/api/sessions', (req, res) => {
  reloadDirectoryCacheIfNeeded();
  const { classCode, department, section, year, teacherName, sessionType, googleSheetUrl } = req.body;

  if (!classCode || !department || !section || !year) {
    return res.status(400).json({ error: "Missing required classroom fields" });
  }

  const normalizedClassCode = classCode.toString().toUpperCase().trim();

  // Helper to isolate sessions between multiple concurrent teachers
  const creatorUsername = req.body.creatorUsername || "admin";

  // Check if teacher sent custom manual geofence coordinates
  let geofence = req.body.geofence;
  if (!geofence || typeof geofence.latitude !== 'number' || typeof geofence.longitude !== 'number') {
    // Find classroom in the synchronized database to load its geofence coordinates
    const room = CLASSROOMS_DIRECTORY.find(
      r => r.classCode.toString().toUpperCase().trim() === normalizedClassCode
    );

    geofence = room ? {
      latitude: parseFloat(room.latitude),
      longitude: parseFloat(room.longitude),
      radius: parseFloat(room.radius) || 30
    } : {
      latitude: 12.9602, // Default Sri Sairam Engineering College center
      longitude: 80.0570,
      radius: 500 // Default campus geofence radius
    };
  }

  sessions[normalizedClassCode] = {
    id: normalizedClassCode, // Classroom code serves as the primary session ID (e.g. G3103)
    creatorUsername: creatorUsername.toLowerCase().trim(),
    classroomCode: normalizedClassCode,
    department,
    section,
    year,
    teacherName: teacherName || "ERP Administrator",
    sessionType: sessionType || "IN",
    geofence,
    googleSheetUrl: googleSheetUrl || "",
    currentOtp: "",
    previousOtps: [],
    currentToken: "",
    previousTokens: [],
    otpUpdatedAt: null,
    students: [],           // { regNo, rollNo, name, timestamp, deviceStatus, method }
    deviceFingerprints: {}, // fingerprint -> [regNo1, regNo2, ...]
    sharingRequests: []     // { regNo, rollNo, name, fingerprint, originalRegNo, timestamp }
  };

  // Set the first OTP and start rotating
  rotateSessionOtp(normalizedClassCode);
  startSessionInterval(normalizedClassCode);

  res.json({ sessionId: normalizedClassCode, ...sessions[normalizedClassCode], intervalId: undefined });
});

// Helper validation for instructor session isolation
function verifySessionOwnership(req, res, session) {
  const teacherUsername = req.headers['x-teacher-username'] || req.query.teacherUsername;
  if (!teacherUsername) {
    res.status(401).json({ error: "Instructor identification header (x-teacher-username) is required." });
    return false;
  }
  if (session.creatorUsername && session.creatorUsername.toLowerCase() !== teacherUsername.toString().toLowerCase().trim()) {
    res.status(403).json({ error: "Access Denied. This session was created by another instructor." });
    return false;
  }
  return true;
}

// Get session state (Teacher Dashboard)
app.get('/api/sessions/:id', (req, res) => {
  const session = sessions[req.params.id];
  if (!session) return res.status(404).json({ error: "Session not found" });

  if (!verifySessionOwnership(req, res, session)) return;

  // Return full details including student list and pending approvals (exclude intervalId)
  const { intervalId, ...rest } = session;
  res.json(rest);
});

// Helper for background auto-sync to Google Sheets in real-time
async function syncSessionToSheetsInBackground(session) {
  if (!session.googleSheetUrl) return;

  try {
    const dateObj = new Date();
    const day = String(dateObj.getDate()).padStart(2, '0');
    const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    const month = months[dateObj.getMonth()];
    const year = dateObj.getFullYear();
    const today = `${day}-${month}-${year}`; // E.g. "01-Jul-2026"

    // Filter students from the roster database that belong to this class
    const classStudents = COLLEGE_STUDENTS_DIRECTORY.filter(s => {
      const deptMatch = String(s.dept).toUpperCase().trim() === String(session.department).toUpperCase().trim();
      const secMatch = String(s.sec).toUpperCase().trim() === String(session.section).toUpperCase().trim();
      const yearMatch = String(s.year).toUpperCase().trim() === String(session.year).toUpperCase().trim();
      return deptMatch && secMatch && yearMatch;
    });

    // Sort students alphabetically by name to maintain a fixed position
    classStudents.sort((a, b) => String(a.name).localeCompare(String(b.name)));

    let studentPayload = [];
    if (classStudents.length > 0) {
      studentPayload = classStudents.map((s, index) => {
        const checkedInList = session.students.filter(
          st => cleanRegNo(st.regNo) === cleanRegNo(s.regNo)
        );
        const presentIn = checkedInList.some(st => !st.type || st.type === "IN");
        const presentOut = checkedInList.some(st => st.type === "OUT");
        return {
          sNo: index + 1,
          regNo: s.regNo.toUpperCase().trim(),
          name: s.name.trim(),
          presentIn: presentIn,
          presentOut: presentOut
        };
      });
    } else {
      studentPayload = session.students.map((s, index) => ({
        sNo: index + 1,
        regNo: s.regNo.toUpperCase().trim(),
        name: s.name.trim(),
        presentIn: !s.type || s.type === "IN",
        presentOut: s.type === "OUT"
      }));
    }

    const payload = {
      metadata: {
        title: "SRI SAIRAM ENGINEERING COLLEGE - CLASS ATTENDANCE REPORT",
        classroomCode: session.classroomCode.toUpperCase(),
        department: session.department,
        section: session.section,
        year: session.year,
        teacherName: session.teacherName || "ERP Administrator",
        date: today,
        sessionType: session.sessionType || "IN",
        classCode: session.id
      },
      students: studentPayload
    };

    await fetch(session.googleSheetUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    });
    console.log(`[Auto-Sync] Synchronized session ${session.id} student database to sheet.`);
  } catch (error) {
    console.error("[Auto-Sync Error]:", error.message);
  }
}

// Get session OTP status (Student verification page)
app.get('/api/sessions/:id/otp-status', (req, res) => {
  const session = sessions[req.params.id.toString().toUpperCase().trim()];
  if (!session) return res.status(404).json({ error: "Session not found" });

  res.json({
    classroomCode: session.classroomCode,
    department: session.department,
    section: session.section,
    year: session.year,
    geofence: session.geofence,
    otpAgeMs: Date.now() - session.otpUpdatedAt
  });
});

// Student checks their device exception approval status (Public)
app.get('/api/sessions/:id/approval-status/:regNo', (req, res) => {
  const session = sessions[req.params.id.toString().toUpperCase().trim()];
  if (!session) return res.status(404).json({ error: "Session not found" });

  const regNo = cleanRegNo(req.params.regNo);
  
  // Check if student has been marked present
  const isMarked = session.students.some(s => cleanRegNo(s.regNo) === regNo);
  if (isMarked) {
    return res.json({ approved: true, pending: false });
  }

  // Check if student is still in pending queue
  const isPending = session.sharingRequests.some(r => cleanRegNo(r.regNo) === regNo);
  return res.json({ approved: false, pending: isPending });
});

// Mark student attendance (Student)
app.post('/api/sessions/:id/attend', (req, res) => {
  reloadDirectoryCacheIfNeeded();
  const session = sessions[req.params.id];
  if (!session) return res.status(404).json({ error: "Session not found" });

  const { regNo, name, fingerprint, otp, token } = req.body;
  if (!regNo || !fingerprint) {
    return res.status(400).json({ error: "Student ID and Fingerprint are required" });
  }

  const regNoLower = cleanRegNo(regNo);

  // 1. Check if student is already marked present for the active session type
  const alreadyMarked = session.students.find(
    s => cleanRegNo(s.regNo) === regNoLower && s.type === (session.sessionType || "IN")
  );
  if (alreadyMarked) {
    return res.status(400).json({ error: `Student is already marked present for the ${session.sessionType || "IN"} session` });
  }

  // 2. Validate student is in official College directory
  const studentInfo = COLLEGE_STUDENTS_DIRECTORY.find(
    s => cleanRegNo(s.regNo) === regNoLower
  );

  if (!studentInfo) {
    return res.status(400).json({ 
      error: "Authentication failed. Student ID not found in official Sairam directory." 
    });
  }

  // 3. Compute geofence range verification (Haversine formula check)
  const coords = req.body.coords || req.body.location;
  if (!coords || (typeof coords.lat !== 'number' && typeof coords.latitude !== 'number') || (typeof coords.lng !== 'number' && typeof coords.longitude !== 'number')) {
    return res.status(400).json({ error: "Location coordinates are required for campus geofence validation" });
  }

  const lat = typeof coords.lat === 'number' ? coords.lat : coords.latitude;
  const lng = typeof coords.lng === 'number' ? coords.lng : coords.longitude;

  const distance = getHaversineDistance(
    lat, 
    lng, 
    session.geofence.latitude, 
    session.geofence.longitude
  );

  if (distance > session.geofence.radius) {
    return res.status(400).json({ 
      error: `Geofencing failure. You are outside the Sairam College geofence area. (Current distance: ${Math.round(distance)}m, Limit: ${session.geofence.radius}m)` 
    });
  }

  // 4. OTP / Token validation (must match current OR previous rotation for network delay grace period)
  const isOtpValid = otp && (otp === session.currentOtp || session.previousOtps.includes(otp));
  const isTokenValid = token && (token === session.currentToken || session.previousTokens.includes(token));

  if (!isOtpValid && !isTokenValid) {
    return res.status(400).json({ 
      error: "Authentication failed. Expired QR code or invalid OTP code. Please look at the Smart TV screen for the active details." 
    });
  }

  // 5. Anti-Cheat: Browser Fingerprint Validation
  if (!session.deviceFingerprints[fingerprint]) {
    session.deviceFingerprints[fingerprint] = [];
  }

  const studentsOnThisDevice = session.deviceFingerprints[fingerprint];

  // If this device has already been used to mark another student's attendance in this session:
  if (studentsOnThisDevice.length > 0 && !studentsOnThisDevice.map(cleanRegNo).includes(cleanRegNo(studentInfo.regNo))) {
    // Check if there is already an approved exception for this student for the active session type
    const isApprovedException = session.students.find(
      s => cleanRegNo(s.regNo) === cleanRegNo(studentInfo.regNo) && 
           s.deviceStatus === "Shared (Approved)" && 
           s.type === (session.sessionType || "IN")
    );

    if (!isApprovedException) {
      // Check if this student is already in the pending approval queue for the active session type
      const alreadyPending = session.sharingRequests.find(
        r => cleanRegNo(r.regNo) === cleanRegNo(studentInfo.regNo) && r.type === (session.sessionType || "IN")
      );

      if (!alreadyPending) {
        session.sharingRequests.push({
          regNo: studentInfo.regNo,
          rollNo: "",
          name: studentInfo.name,
          type: session.sessionType || "IN",
          fingerprint,
          originalRegNo: studentsOnThisDevice[0], // Who marked first on this phone
          timestamp: new Date().toLocaleTimeString()
        });
      }

      return res.status(202).json({
        status: "pending_approval",
        message: `This device was already used by another student. Sent a device sharing request to your Teacher. Please ask them to approve your request on the Smart TV screen.`
      });
    }
  }

  // 6. Perfect validation - Mark attendance
  const normalizedReg = cleanRegNo(studentInfo.regNo);
  if (!studentsOnThisDevice.map(cleanRegNo).includes(normalizedReg)) {
    studentsOnThisDevice.push(normalizedReg);
  }
  session.deviceFingerprints[fingerprint] = studentsOnThisDevice;

  const markedStudent = {
    regNo: studentInfo.regNo,
    name: studentInfo.name,
    type: session.sessionType || "IN",
    timestamp: new Date().toLocaleTimeString(),
    deviceStatus: studentsOnThisDevice.length > 1 ? "Shared (Direct)" : "Single Device",
    method: otp ? "OTP Code" : "QR Scan"
  };

  session.students.push(markedStudent);

  res.json({
    status: "success",
    message: "Attendance marked successfully!",
    student: markedStudent
  });
});

// Teacher approves pending device sharing request
app.post('/api/sessions/:id/approve-sharing', (req, res) => {
  const session = sessions[req.params.id];
  if (!session) return res.status(404).json({ error: "Session not found" });

  if (!verifySessionOwnership(req, res, session)) return;

  const { regNo } = req.body;
  if (!regNo) return res.status(400).json({ error: "Missing Register Number" });

  // Find index in pending queue
  const requestIndex = session.sharingRequests.findIndex(
    r => cleanRegNo(r.regNo) === cleanRegNo(regNo)
  );

  if (requestIndex === -1) {
    return res.status(404).json({ error: "Pending request not found" });
  }

  const request = session.sharingRequests[requestIndex];

  // Remove from pending queue
  session.sharingRequests.splice(requestIndex, 1);

  // Bind fingerprint to this student as well
  if (!session.deviceFingerprints[request.fingerprint].includes(request.regNo)) {
    session.deviceFingerprints[request.fingerprint].push(request.regNo);
  }

  // Mark present
  const markedStudent = {
    regNo: request.regNo,
    name: request.name,
    type: session.sessionType || "IN",
    timestamp: new Date().toLocaleTimeString(),
    deviceStatus: "Shared (Approved)",
    method: "Teacher Approved"
  };

  session.students.push(markedStudent);

  res.json({ status: "success", message: `Approved device sharing for ${request.name}`, student: markedStudent });
});

// Teacher rejects pending device sharing request
app.post('/api/sessions/:id/reject-sharing', (req, res) => {
  const session = sessions[req.params.id];
  if (!session) return res.status(404).json({ error: "Session not found" });

  if (!verifySessionOwnership(req, res, session)) return;

  const { regNo } = req.body;
  if (!regNo) return res.status(400).json({ error: "Missing Register Number" });

  const requestIndex = session.sharingRequests.findIndex(
    r => cleanRegNo(r.regNo) === cleanRegNo(regNo)
  );

  if (requestIndex === -1) {
    return res.status(404).json({ error: "Pending request not found" });
  }

  session.sharingRequests.splice(requestIndex, 1);
  res.json({ status: "success", message: "Rejected device sharing request" });
});

// Teacher manually marks student present (e.g. if everything fails)
app.post('/api/sessions/:id/manual-mark', (req, res) => {
  const session = sessions[req.params.id];
  if (!session) return res.status(404).json({ error: "Session not found" });

  if (!verifySessionOwnership(req, res, session)) return;

  const { regNo } = req.body;
  if (!regNo) return res.status(400).json({ error: "Missing Register Number" });

  // Check if student is in college directory
  const directoryMatch = COLLEGE_STUDENTS_DIRECTORY.find(
    s => cleanRegNo(s.regNo) === cleanRegNo(regNo)
  );

  // If already marked for the active session type, fail
  const alreadyMarked = session.students.find(
    s => cleanRegNo(s.regNo) === cleanRegNo(regNo) && s.type === (session.sessionType || "IN")
  );
  if (alreadyMarked) {
    return res.status(400).json({ error: `Student already marked present for the ${session.sessionType || "IN"} session` });
  }

  const regNoNormalized = regNo.toUpperCase().trim();
  const markedStudent = {
    regNo: regNoNormalized,
    name: directoryMatch ? directoryMatch.name : "Unknown Student",
    type: session.sessionType || "IN",
    timestamp: new Date().toLocaleTimeString(),
    deviceStatus: "Manually Added",
    method: "Teacher Manual"
  };

  session.students.push(markedStudent);

  res.json({ status: "success", message: `Successfully marked ${markedStudent.name} present manually`, student: markedStudent });
});

// Teacher edits a student's check-in details (e.g. if mistake made)
app.post('/api/sessions/:id/edit-student', (req, res) => {
  const session = sessions[req.params.id];
  if (!session) return res.status(404).json({ error: "Session not found" });

  if (!verifySessionOwnership(req, res, session)) return;

  const { oldRegNo, regNo, name } = req.body;
  if (!oldRegNo || !regNo || !name) {
    return res.status(400).json({ error: "Missing required student parameters" });
  }

  const student = session.students.find(
    s => cleanRegNo(s.regNo) === cleanRegNo(oldRegNo)
  );

  if (!student) {
    return res.status(404).json({ error: "Student not found in active session attendance" });
  }

  student.regNo = regNo.toUpperCase().trim();
  student.name = name.trim();

  for (const fp in session.deviceFingerprints) {
    const list = session.deviceFingerprints[fp];
    const index = list.map(cleanRegNo).indexOf(cleanRegNo(oldRegNo));
    if (index !== -1) {
      list[index] = cleanRegNo(regNo);
    }
  }

  res.json({ 
    status: "success", 
    message: "Student details updated successfully", 
    student 
  });
});

// Sync data to Google Sheet via the Google Apps Script Webhook (Structured layout)
app.post('/api/sessions/:id/sync-sheets', async (req, res) => {
  const session = sessions[req.params.id];
  if (!session) return res.status(404).json({ error: "Session not found" });

  if (!verifySessionOwnership(req, res, session)) return;

  const sheetUrl = session.googleSheetUrl || req.body.googleSheetUrl;
  if (!sheetUrl) {
    return res.status(400).json({ error: "Google Apps Script Web App URL is not configured. Please add one in settings." });
  }

  try {
    // Send attendance data formatted into a clean structured report payload (strictly DD-MMM-YYYY text)
    const dateObj = new Date();
    const day = String(dateObj.getDate()).padStart(2, '0');
    const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    const month = months[dateObj.getMonth()];
    const year = dateObj.getFullYear();
    const today = `${day}-${month}-${year}`; // E.g. "01-Jul-2026"

    // Filter students from the roster database that belong to this class
    const classStudents = COLLEGE_STUDENTS_DIRECTORY.filter(s => {
      const deptMatch = String(s.dept).toUpperCase().trim() === String(session.department).toUpperCase().trim();
      const secMatch = String(s.sec).toUpperCase().trim() === String(session.section).toUpperCase().trim();
      const yearMatch = String(s.year).toUpperCase().trim() === String(session.year).toUpperCase().trim();
      return deptMatch && secMatch && yearMatch;
    });

    // Sort students alphabetically by name to maintain a fixed position
    classStudents.sort((a, b) => String(a.name).localeCompare(String(b.name)));

    let studentPayload = [];
    if (classStudents.length > 0) {
      studentPayload = classStudents.map((s, index) => {
        const checkedInList = session.students.filter(
          st => cleanRegNo(st.regNo) === cleanRegNo(s.regNo)
        );
        const presentIn = checkedInList.some(st => !st.type || st.type === "IN");
        const presentOut = checkedInList.some(st => st.type === "OUT");
        return {
          sNo: index + 1,
          regNo: s.regNo.toUpperCase().trim(),
          name: s.name.trim(),
          presentIn: presentIn,
          presentOut: presentOut
        };
      });
    } else {
      // Fallback: if database roster is empty, send present students
      studentPayload = session.students.map((s, index) => ({
        sNo: index + 1,
        regNo: s.regNo.toUpperCase().trim(),
        name: s.name.trim(),
        presentIn: !s.type || s.type === "IN",
        presentOut: s.type === "OUT"
      }));
    }

    const payload = {
      metadata: {
        title: "SRI SAIRAM ENGINEERING COLLEGE - CLASS ATTENDANCE REPORT",
        classroomCode: session.classroomCode.toUpperCase(),
        department: session.department,
        section: session.section,
        year: session.year,
        teacherName: session.teacherName || "ERP Administrator",
        date: today,
        sessionType: session.sessionType || "IN",
        classCode: session.id
      },
      students: studentPayload
    };

    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 6500); // 6.5s timeout guard

    const response = await fetch(sheetUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify(payload),
      signal: controller.signal,
      mode: "cors"
    });
    clearTimeout(timeoutId);

    const resultText = await response.text();
    let result = {};
    try {
      result = JSON.parse(resultText);
    } catch (e) {
      result = { raw: resultText };
    }

    res.json({ status: "success", message: "Synchronized with Google Sheets successfully!", details: result });
  } catch (error) {
    console.error("Sheets sync failed:", error);
    res.status(500).json({ error: `Connection to Google Apps Script failed: ${error.message}` });
  }
});

// Toggle Session Type (Teacher active dashboard - IN vs OUT)
app.put('/api/sessions/:id/session-type', (req, res) => {
  const session = sessions[req.params.id];
  if (!session) return res.status(404).json({ error: "Session not found" });

  if (!verifySessionOwnership(req, res, session)) return;

  const { sessionType } = req.body;
  if (sessionType !== 'IN' && sessionType !== 'OUT') {
    return res.status(400).json({ error: "Invalid session type. Must be 'IN' or 'OUT'" });
  }

  session.sessionType = sessionType;
  
  // Rotate token immediately to reset OTP and refresh scanner parameters
  rotateSessionOtp(session.id);
  
  res.json({ success: true, sessionType: session.sessionType });
});

// Delete session / Clear
app.delete('/api/sessions/:id', (req, res) => {
  const session = sessions[req.params.id];
  if (!session) return res.status(404).json({ error: "Session not found" });

  if (!verifySessionOwnership(req, res, session)) return;

  if (session.intervalId) {
    clearInterval(session.intervalId);
  }
  delete sessions[req.params.id];
  res.json({ status: "success", message: "Session closed successfully" });
});

// Check status of in-memory roster database (Self-Healing helper)
app.get('/api/database/status', (req, res) => {
  reloadDirectoryCacheIfNeeded();
  res.json({
    empty: COLLEGE_STUDENTS_DIRECTORY.length === 0,
    facultyCount: FACULTY_DIRECTORY.length,
    studentCount: COLLEGE_STUDENTS_DIRECTORY.length,
    classroomCount: CLASSROOMS_DIRECTORY.length
  });
});

// Restore roster database (Self-Healing recovery protocol)
app.post('/api/database/restore', (req, res) => {
  const { faculty, students, classrooms } = req.body;
  
  if (Array.isArray(faculty) && Array.isArray(students)) {
    FACULTY_DIRECTORY = faculty;
    COLLEGE_STUDENTS_DIRECTORY = students;
    if (Array.isArray(classrooms)) {
      CLASSROOMS_DIRECTORY = classrooms;
    }
    
    // Write back to local cache file
    try {
      fs.writeFileSync(
        cachePath,
        JSON.stringify({ faculty: FACULTY_DIRECTORY, students: COLLEGE_STUDENTS_DIRECTORY, classrooms: CLASSROOMS_DIRECTORY }, null, 2)
      );
      console.log("[Self-Healing] Restored roster databases synced to local cache file.");
    } catch (err) {
      console.error("Self-healing cache write failed:", err.message);
    }
    
    return res.json({ success: true, message: "Database restored successfully!" });
  }
  
  res.status(400).json({ error: "Invalid database recovery payload format" });
});

// Retrieve Classrooms database details (Used in Teacher setup dropdowns)
app.get('/api/database/classrooms', (req, res) => {
  reloadDirectoryCacheIfNeeded();
  res.json(CLASSROOMS_DIRECTORY);
});

// Dynamically synchronize the College Faculty & Student Roster database from Google Sheets
app.post('/api/database/sync-roster', async (req, res) => {
  const { rosterDbUrl } = req.body;
  console.log("[Roster Sync Debug] Sync request received for URL:", rosterDbUrl);
  if (!rosterDbUrl) {
    return res.status(400).json({ error: "Missing Roster Web App URL parameter" });
  }

  try {
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 6500); // 6.5s timeout guard
    
    // Automatically append the security sync authorization key parameter
    const syncKey = process.env.ROSTER_SYNC_KEY || "sairamsynckey2026";
    const separator = rosterDbUrl.includes('?') ? '&' : '?';
    const authorizedUrl = `${rosterDbUrl}${separator}key=${syncKey}`;
    
    const response = await fetch(authorizedUrl, { signal: controller.signal });
    clearTimeout(timeoutId);
    
    const text = await response.text();
    console.log("[Roster Sync Debug] Raw sheet response length:", text.length);
    
    let data;
    try {
      data = JSON.parse(text);
    } catch (parseErr) {
      console.error("[Roster Sync Debug] JSON parse failed. Raw response snippet:", text.substring(0, 300));
      return res.status(400).json({ error: "Roster Sheet returned invalid JSON. Verify Web App deployment." });
    }

    console.log("[Roster Sync Debug] Parsed keys:", Object.keys(data));
    if (data && data.status === "success") {
      const faculty = data.faculty || [];
      const students = data.students || [];
      
      console.log(`[Roster Sync Debug] Retrieved: ${faculty.length} faculty, ${students.length} students.`);

      if (faculty.length === 0 || students.length === 0) {
        return res.status(400).json({ 
          error: "Synchronization failed. Your Google Sheet 'Students' or 'Faculty' sheet tab appears to be empty." 
        });
      }

      FACULTY_DIRECTORY = faculty;
      COLLEGE_STUDENTS_DIRECTORY = students;
      
      // Handle both classrooms and old courses formats
      if (Array.isArray(data.classrooms)) {
        CLASSROOMS_DIRECTORY = data.classrooms;
      } else if (Array.isArray(data.courses)) {
        CLASSROOMS_DIRECTORY = data.courses.map(c => ({
          classCode: c.courseCode || c.code || "UNKNOWN",
          latitude: "12.960200",
          longitude: "80.057000",
          radius: "500"
        }));
      }

      // Cache synced directory locally on disk
      try {
        fs.writeFileSync(
          cachePath,
          JSON.stringify({ faculty: FACULTY_DIRECTORY, students: COLLEGE_STUDENTS_DIRECTORY, classrooms: CLASSROOMS_DIRECTORY }, null, 2)
        );
        console.log("[DB Cache] Synced roster written to disk cache.");
      } catch (err) {
        console.error("Failed to write roster disk cache:", err.message);
      }
      
      console.log(`[DB Sync] Loaded ${FACULTY_DIRECTORY.length} faculty, ${COLLEGE_STUDENTS_DIRECTORY.length} students, and ${CLASSROOMS_DIRECTORY.length} classrooms from sheet.`);
      
      res.json({
        status: "success",
        message: "Successfully synchronized college database!",
        facultyCount: FACULTY_DIRECTORY.length,
        studentCount: COLLEGE_STUDENTS_DIRECTORY.length,
        classroomCount: CLASSROOMS_DIRECTORY.length,
        faculty: FACULTY_DIRECTORY,
        students: COLLEGE_STUDENTS_DIRECTORY,
        classrooms: CLASSROOMS_DIRECTORY
      });
    } else {
      res.status(400).json({ 
        error: "Invalid roster schema. Web App returned status: " + (data ? data.status : "undefined")
      });
    }
  } catch (error) {
    console.error("Roster DB Sync Failed:", error);
    res.status(500).json({ error: `Connection to Roster Sheet failed: ${error.message}` });
  }
});

// Start dev server boot (omitted on Vercel runtime context)
if (process.env.NODE_ENV !== 'production' || !process.env.VERCEL) {
  app.listen(PORT, () => {
    console.log(`[ERP Server] Sairam College Attendance Server running on port ${PORT}`);
  });
}

export default app;
