import React, { useState, useEffect, useRef } from 'react';
import { 
  QrCode, Tv, CheckCircle2, MapPin, AlertTriangle, Loader2, 
  FileSpreadsheet, Users, X, Settings, Camera,
  ChevronRight, Plus, RefreshCw, HelpCircle, Send, Smartphone, Lock, Clipboard, Edit2
} from 'lucide-react';
import { QRCodeSVG } from 'qrcode.react';
import { Html5Qrcode } from 'html5-qrcode';
import { verifyGeofence, SAIRAM_CAMPUS_COORDS } from './utils/geo';
import { getDeviceFingerprint, getDeviceDescription } from './utils/fingerprint';

const APPS_SCRIPT_DOGET = `function doGet(e) {
  // Security Access Authorization Check (Prevents public data exposure)
  var SYNC_KEY = "sairamsynckey2026"; // Feel free to customize this key!
  var requestKey = e && e.parameter && e.parameter.key ? e.parameter.key.toString().trim() : "";
  if (requestKey !== SYNC_KEY) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Unauthorized access. Invalid sync key." }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    var facultySheet = ss.getSheetByName("Faculty");
    var facultyData = [];
    if (facultySheet) {
      var facultyRows = facultySheet.getDataRange().getValues();
      var headers = facultyRows[0].map(function(h) { return h.toString().toLowerCase().trim(); });
      
      for (var i = 1; i < facultyRows.length; i++) {
        var row = facultyRows[i];
        var item = {};
        for (var j = 0; j < headers.length; j++) {
          var key = headers[j];
          if (key === "username" || key === "faculty id") item.username = row[j].toString().toLowerCase().trim();
          else if (key === "password") item.password = row[j].toString().trim();
          else if (key === "name" || key === "faculty name") item.name = row[j].toString().trim();
          else if (key === "department" || key === "dept") item.dept = row[j].toString().trim();
        }
        if (item.username && item.password) facultyData.push(item);
      }
    }
    
    var studentSheet = ss.getSheetByName("Students");
    var studentData = [];
    if (studentSheet) {
      var studentRows = studentSheet.getDataRange().getValues();
      var headers = studentRows[0].map(function(h) { return h.toString().toLowerCase().trim(); });
      
      for (var i = 1; i < studentRows.length; i++) {
        var row = studentRows[i];
        var item = {};
        for (var j = 0; j < headers.length; j++) {
          var key = headers[j];
          if (key === "student id" || key === "id") {
            item.regNo = row[j].toString().replace(/[^a-zA-Z0-9]/g, '').toLowerCase().trim();
          } else if (key === "student name" || key === "name") {
            item.name = row[j].toString().trim();
          } else if (key === "department" || key === "dept") {
            item.dept = row[j].toString().trim();
          } else if (key === "section" || key === "sec") {
            item.sec = row[j].toString().trim();
          } else if (key === "year") {
            item.year = row[j].toString().trim();
          }
        }
        if (item.regNo) studentData.push(item);
      }
    }

    var classroomsSheet = ss.getSheetByName("Classrooms");
    var classroomsData = [];
    if (classroomsSheet) {
      var classroomsRows = classroomsSheet.getDataRange().getValues();
      var headers = classroomsRows[0].map(function(h) { return h.toString().toLowerCase().trim(); });
      
      for (var i = 1; i < classroomsRows.length; i++) {
        var row = classroomsRows[i];
        var item = {};
        for (var j = 0; j < headers.length; j++) {
          var key = headers[j];
          if (key === "class code" || key === "classroom code" || key === "code") {
            item.classCode = row[j].toString().toUpperCase().trim();
          } else if (key === "latitude" || key === "lat") {
            item.latitude = row[j].toString().trim();
          } else if (key === "longitude" || key === "lng") {
            item.longitude = row[j].toString().trim();
          } else if (key === "radius") {
            item.radius = row[j].toString().trim();
          }
        }
        if (item.classCode && item.latitude && item.longitude) classroomsData.push(item);
      }
    }
    
    return ContentService.createTextOutput(JSON.stringify({ status: "success", faculty: facultyData, students: studentData, classrooms: classroomsData }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
`;

const APPS_SCRIPT_DOPOST = `function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = JSON.parse(e.postData.contents);
    
    var isStructured = data && data.metadata && data.students;
    
    if (isStructured) {
      // 1. Initialize Sheet headers if it's a fresh/empty sheet
      if (sheet.getLastRow() < 6) {
        sheet.clear();
        
        // Title Block
        sheet.getRange("A1").setValue(data.metadata.title);
        sheet.getRange("A1").setFontWeight("bold").setFontSize(14);
        sheet.getRange("A1:G1").merge().setHorizontalAlignment("center");
        
        // Metadata grid settings
        sheet.getRange("A2").setValue("Classroom Code:").setFontWeight("bold");
        sheet.getRange("B2").setValue(data.metadata.classroomCode);
        sheet.getRange("C2").setValue("Dept: " + data.metadata.department).setFontWeight("bold");
        
        sheet.getRange("A3").setValue("Year:").setFontWeight("bold");
        sheet.getRange("B3").setValue(data.metadata.year);
        sheet.getRange("C3").setValue("Section: " + data.metadata.section).setFontWeight("bold");

        sheet.getRange("A4").setValue("Faculty In-Charge:").setFontWeight("bold");
        sheet.getRange("B4").setValue(data.metadata.teacherName);
        
        // Roster Student Details Columns (No Roll Number column!)
        sheet.getRange("A6").setValue("S.No").setFontWeight("bold").setHorizontalAlignment("center").setBackground("#0f4c81").setFontColor("#ffffff");
        sheet.getRange("B6").setValue("Student ID").setFontWeight("bold").setHorizontalAlignment("center").setBackground("#0f4c81").setFontColor("#ffffff");
        sheet.getRange("C6").setValue("Student Name").setFontWeight("bold").setHorizontalAlignment("center").setBackground("#0f4c81").setFontColor("#ffffff");
        
        sheet.setColumnWidth(1, 45);  // S.No
        sheet.setColumnWidth(2, 120); // Student ID
        sheet.setColumnWidth(3, 160); // Student Name
      }
      
      // 2. Sync Roster Student listings (Add any missing students, sort alphabetically)
      var lastRow = sheet.getLastRow();
      var studentRows = [];
      if (lastRow >= 7) {
        var range = sheet.getRange(7, 2, lastRow - 6, 2); // Column B & C
        var vals = range.getValues();
        for (var k = 0; k < vals.length; k++) {
          studentRows.push({
            rowNum: 7 + k,
            regNo: vals[k][0].toString().replace(/[^a-zA-Z0-9]/g, '').toLowerCase().trim(),
            name: vals[k][1].toString().trim()
          });
        }
      }

      if (studentRows.length === 0) {
        // First population: write the entire class roster from incoming data
        for (var i = 0; i < data.students.length; i++) {
          var st = data.students[i];
          sheet.appendRow([
            i + 1,
            st.regNo.toUpperCase().trim(),
            st.name.trim()
          ]);
        }
        
        // Reload list mapping
        lastRow = sheet.getLastRow();
        var range = sheet.getRange(7, 2, lastRow - 6, 2);
        var vals = range.getValues();
        for (var k = 0; k < vals.length; k++) {
          studentRows.push({
            rowNum: 7 + k,
            regNo: vals[k][0].toString().replace(/[^a-zA-Z0-9]/g, '').toLowerCase().trim(),
            name: vals[k][1].toString().trim()
          });
        }
        // Look for any newly synced/missing student registry row
        var sheetChanged = false;
        for (var i = 0; i < data.students.length; i++) {
          var st = data.students[i];
          var exists = studentRows.some(function(r) { return r.regNo.toLowerCase().replace(/[^a-zA-Z0-9]/g, '') === st.regNo.toLowerCase().replace(/[^a-zA-Z0-9]/g, ''); });
          if (!exists) {
            sheet.appendRow([
              studentRows.length + 1,
              st.regNo.toUpperCase().trim(),
              st.name.trim()
            ]);
            studentRows.push({
              rowNum: 7 + studentRows.length,
              regNo: st.regNo.toLowerCase().replace(/[^a-zA-Z0-9]/g, ''),
              name: st.name.trim()
            });
            sheetChanged = true;
          }
        }
        
        if (sheetChanged) {
          lastRow = sheet.getLastRow();
          
          // Sort entire table rows starting from row 7 alphabetically by Name (Column C)
          var sortRange = sheet.getRange(7, 1, lastRow - 6, sheet.getLastColumn());
          sortRange.sort({ column: 3, ascending: true });
          
          // Re-write the S.No index numbers consecutively
          var sNoRange = sheet.getRange(7, 1, lastRow - 6, 1);
          var sNoVals = [];
          for (var n = 1; n <= (lastRow - 6); n++) {
            sNoVals.push([n]);
          }
          sNoRange.setValues(sNoVals);
          
          // Refresh list mapping
          studentRows = [];
          var range = sheet.getRange(7, 2, lastRow - 6, 2);
          var vals = range.getValues();
          for (var k = 0; k < vals.length; k++) {
            studentRows.push({
              rowNum: 7 + k,
              regNo: vals[k][0].toString().replace(/[^a-zA-Z0-9]/g, '').toLowerCase().trim(),
              name: vals[k][1].toString().trim()
            });
          }
        }
      }
      
      // 3. Find or Create date columns
      var dateStr = data.metadata.date;
      var lastCol = sheet.getLastColumn();
      var dateColIndex = -1;
      
      if (lastCol >= 4) {
        var dateRowValues = sheet.getRange(5, 4, 1, Math.max(1, lastCol - 3)).getValues()[0];
        for (var c = 0; c < dateRowValues.length; c += 2) {
          var cellVal = dateRowValues[c];
          var cellStr = "";
          if (cellVal instanceof Date) {
            cellStr = Utilities.formatDate(cellVal, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || "GMT+5:30", "dd-MMM-yyyy");
          } else {
            cellStr = cellVal ? cellVal.toString().trim() : "";
          }
          if (cellStr.toLowerCase() === dateStr.toLowerCase().trim()) {
            dateColIndex = 4 + c;
            break;
          }
        }
      }
      
      if (dateColIndex === -1) {
        // Date not found, create new merged headers (E.g. Column D & E for today's entry)
        dateColIndex = Math.max(4, lastCol + 1);
        sheet.getRange(5, dateColIndex, 1, 2).merge()
          .setValue("'" + dateStr) // Prefix with single quote to force string type
          .setFontWeight("bold")
          .setHorizontalAlignment("center")
          .setBackground("#e2e8f0")
          .setBorder(true, true, true, true, true, true);
          
        sheet.getRange(6, dateColIndex).setValue("IN").setFontWeight("bold").setHorizontalAlignment("center").setBorder(true, true, true, true, true, true);
        sheet.getRange(6, dateColIndex + 1).setValue("OUT").setFontWeight("bold").setHorizontalAlignment("center").setBorder(true, true, true, true, true, true);
        
        sheet.setColumnWidth(dateColIndex, 60);
        sheet.setColumnWidth(dateColIndex + 1, 60);
      }
      
      // 4. Mark present (P in Green) or absent (A in Red) for each row based on IN/OUT type
      var isOutPhase = (data.metadata.sessionType === "OUT");
      var inCol = dateColIndex;
      var outCol = dateColIndex + 1;
      
      for (var r = 0; r < studentRows.length; r++) {
        var student = studentRows[r];
        var studentPayloadItem = data.students.find(function(s) {
          return s.regNo.toLowerCase() === student.regNo;
        });
        
        // --- 4a. Mark IN (Entry) Attendance ---
        var inCell = sheet.getRange(student.rowNum, inCol);
        if (studentPayloadItem && studentPayloadItem.presentIn) {
          inCell.setValue("P")
            .setFontColor("#15803d")   // Green text
            .setBackground("#f0fdf4")  // Green background
            .setFontWeight("bold")
            .setHorizontalAlignment("center")
            .setBorder(true, true, true, true, true, true);
        } else {
          var currentInVal = inCell.getValue().toString().trim();
          if (currentInVal === "" || currentInVal === "A" || isOutPhase) {
            inCell.setValue("A")
              .setFontColor("#b91c1c")   // Red text
              .setBackground("#fef2f2")  // Red background
              .setFontWeight("bold")
              .setHorizontalAlignment("center")
              .setBorder(true, true, true, true, true, true);
          }
        }
        // --- 4b. Mark OUT (Exit) Attendance ---
        var outCell = sheet.getRange(student.rowNum, outCol);
        if (isOutPhase) {
          if (studentPayloadItem && studentPayloadItem.presentOut) {
            outCell.setValue("P")
              .setFontColor("#15803d")   // Green text
              .setBackground("#f0fdf4")  // Green background
              .setFontWeight("bold")
              .setHorizontalAlignment("center")
              .setBorder(true, true, true, true, true, true);
          } else {
            var inVal = inCell.getValue().toString().trim();
            if (inVal === "A" || inVal === "") {
              var currentOutVal = outCell.getValue().toString().trim();
              if (currentOutVal === "" || currentOutVal === "A") {
                outCell.setValue("A")
                  .setFontColor("#b91c1c")   // Red text
                  .setBackground("#fef2f2")  // Red background
                  .setFontWeight("bold")
                  .setHorizontalAlignment("center")
                  .setBorder(true, true, true, true, true, true);
              }
            } else {
              // If present for IN but has not marked OUT, clear it so it stays blank!
              outCell.setValue("")
                .setBackground(null)
                .setFontColor(null)
                .setFontWeight("normal")
                .setBorder(true, true, true, true, true, true);
            }
          }
        }
      }
      
      // 5. Merge the top title cell dynamically across all columns to prevent single-cell truncation
      var finalColCount = sheet.getLastColumn();
      var titleRange = sheet.getRange(1, 1, 1, Math.max(7, finalColCount));
      titleRange.breakApart();
      titleRange.merge().setHorizontalAlignment("center");
    } else {
      // Fallback for raw flat sheet logs
      var headers = ["Student ID", "Student Name", "Timestamp", "Method", "Device Status"];
      if (sheet.getLastRow() === 0) {
        sheet.appendRow(headers);
      }
      var studentsList = Array.isArray(data) ? data : [];
      for (var i = 0; i < studentsList.length; i++) {
        var student = studentsList[i];
        sheet.appendRow([
          student.regNo.toUpperCase(),
          student.name,
          student.timestamp,
          student.method,
          student.deviceStatus
        ]);
      }
    }
  } catch (error) {
    // Silent fail
}
`;

interface StudentAttendance {
  regNo: string;
  rollNo?: string;
  name: string;
  type?: 'IN' | 'OUT';
  timestamp: string;
  deviceStatus: string;
  method: string;
}

interface SharingRequest {
  regNo: string;
  rollNo?: string;
  name: string;
  type?: 'IN' | 'OUT';
  fingerprint: string;
  originalRegNo: string;
  timestamp: string;
}

interface ActiveSession {
  id: string;
  classroomCode: string;
  department: string;
  section: string;
  year: string;
  teacherName: string;
  geofence: {
    latitude: number;
    longitude: number;
    radius: number;
  };
  googleSheetUrl: string;
  currentOtp: string;
  currentToken: string;
  students: StudentAttendance[];
  sharingRequests: SharingRequest[];
  sessionType: 'IN' | 'OUT';
}

interface ClassLookupInfo {
  classroomCode: string;
  department: string;
  section: string;
  year: string;
  teacherName: string;
}

export default function App() {
  const [view, setView] = useState<'landing' | 'teacher_login' | 'teacher_setup' | 'teacher_dashboard' | 'student' | 'student_success'>('landing');
  
  // Teacher states
  const [teacherClassCode, setTeacherClassCode] = useState('');
  const [department, setDepartment] = useState('IT');
  const [section, setSection] = useState('A');
  const [year, setYear] = useState('3rd');
  const [geofenceLat, setGeofenceLat] = useState<string>(SAIRAM_CAMPUS_COORDS.latitude.toString());
  const [geofenceLng, setGeofenceLng] = useState<string>(SAIRAM_CAMPUS_COORDS.longitude.toString());
  const [geofenceRadius, setGeofenceRadius] = useState<string>("30"); // Default 30m classroom bounds
  const [googleSheetUrl, setGoogleSheetUrl] = useState('');
  const [activeSession, setActiveSession] = useState<ActiveSession | null>(null);
  const [teacherCustomIp, setTeacherCustomIp] = useState('');
  const [searchRegNo, setSearchRegNo] = useState('');
  const [showSettings, setShowSettings] = useState(false);
  const [showAppsScriptModal, setShowAppsScriptModal] = useState(false);
  const [syncStatus, setSyncStatus] = useState<{ success?: boolean; message?: string } | null>(null);
  const [isSyncing, setIsSyncing] = useState(false);

  // Student states
  const [classCode, setClassCode] = useState('');
  const [classLookup, setClassLookup] = useState<ClassLookupInfo | null>(null);
  const [isClassSearching, setIsClassSearching] = useState(false);
  const [classLookupError, setClassLookupError] = useState('');

  // Scanner states
  const [showScanner, setShowScanner] = useState(false);
  const [scannerError, setScannerError] = useState('');
  const scannerRef = useRef<Html5Qrcode | null>(null);

  // Camera Zoom Control States
  const [zoomSupported, setZoomSupported] = useState(false);
  const [minZoom, setMinZoom] = useState(1);
  const [maxZoom, setMaxZoom] = useState(4);
  const [currentZoom, setCurrentZoom] = useState(1);

  // Helper to retrieve the active camera media stream video track
  const getCameraTrack = (): MediaStreamTrack | null => {
    const videoElem = document.querySelector("#qr-reader video") as HTMLVideoElement;
    if (videoElem && videoElem.srcObject) {
      const stream = videoElem.srcObject as MediaStream;
      const tracks = stream.getVideoTracks();
      if (tracks.length > 0) {
        return tracks[0];
      }
    }
    return null;
  };

  // Dynamically apply camera zoom constraints
  const handleZoomChange = async (val: number) => {
    setCurrentZoom(val);
    const track = getCameraTrack();
    if (track) {
      try {
        await track.applyConstraints({
          advanced: [{ zoom: val } as any]
        });
      } catch (err) {
        console.warn("Failed to apply camera zoom constraint:", err);
      }
    }
  };

  // Faculty login states
  const [facultyUser, setFacultyUser] = useState('');
  const [facultyPass, setFacultyPass] = useState('');
  const [activeTeacherName, setActiveTeacherName] = useState('');
  const [facultyLoginError, setFacultyLoginError] = useState('');
  const [isLoggingIn, setIsLoggingIn] = useState(false);

  const [studentRegNo, setStudentRegNo] = useState('');
  const [studentOtp, setStudentOtp] = useState('');
  const [scannedToken, setScannedToken] = useState('');
  const [studentName, setStudentName] = useState('');
  const [studentMatchingMsg, setStudentMatchingMsg] = useState('');
  const [geoStatus, setGeoStatus] = useState<'idle' | 'checking' | 'success' | 'out_of_bounds' | 'denied'>('idle');
  const [studentDistance, setStudentDistance] = useState<number | null>(null);
  const [studentCoords, setStudentCoords] = useState<{ lat: number; lng: number } | null>(null);
  const [classGeofence, setClassGeofence] = useState<{ latitude: number; longitude: number; radius: number } | null>(null);
  // Roster Sheet DB Sync States
  const [rosterUrl, setRosterUrl] = useState(() => localStorage.getItem('sairam_roster_url') || '');
  const [isSyncingRoster, setIsSyncingRoster] = useState(false);
  const [rosterSyncMsg, setRosterSyncMsg] = useState('');
  const [showRosterGuideModal, setShowRosterGuideModal] = useState(false);
  const [studentDeviceFingerprint] = useState(() => getDeviceFingerprint());
  const [studentDeviceDesc] = useState(() => getDeviceDescription());
  const [attendanceError, setAttendanceError] = useState('');
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [pendingApproval, setPendingApproval] = useState(false);
  const [markedStudentDetails, setMarkedStudentDetails] = useState<StudentAttendance | null>(null);

  // Dynamic Classrooms states
  const [classroomsList, setClassroomsList] = useState<{ classCode: string; latitude: string; longitude: string; radius: string }[]>([]);
  const [selectedClassroomIndex, setSelectedClassroomIndex] = useState<string>('');
  const [sessionType, setSessionType] = useState<'IN' | 'OUT'>('IN');

  // Password Change states
  const [showChangePasswordModal, setShowChangePasswordModal] = useState(false);
  const [currentPassword, setCurrentPassword] = useState('');
  const [newPassword, setNewPassword] = useState('');
  const [confirmNewPassword, setConfirmNewPassword] = useState('');
  const [changePasswordMsg, setChangePasswordMsg] = useState('');
  const [isChangingPassword, setIsChangingPassword] = useState(false);

  // Edit Student states
  const [showEditStudentModal, setShowEditStudentModal] = useState(false);
  const [editingStudent, setEditingStudent] = useState<StudentAttendance | null>(null);
  const [editStudentName, setEditStudentName] = useState('');
  const [editStudentReg, setEditStudentReg] = useState('');

  // Admin Setup lock states
  const [adminKey, setAdminKey] = useState('');
  const [isAdminUnlocked, setIsAdminUnlocked] = useState(false);

  // OTP Progress Counter
  const [otpProgress, setOtpProgress] = useState(100);

  // Refs for polling
  const pollingRef = useRef<number | null>(null);
  const approvalPollingRef = useRef<number | null>(null);
  const gpsWatchRef = useRef<number | null>(null);

  // Detect query parameters on load (e.g. when QR code is scanned)
  useEffect(() => {
    const params = new URLSearchParams(window.location.search);
    const sessionParam = params.get('session');
    const tokenParam = params.get('token');
    
    if (sessionParam) {
      setClassCode(sessionParam);
      if (tokenParam) {
        setScannedToken(tokenParam);
        setStudentOtp('QR-SCAN'); // Set placeholder indicating scanned scan
      }
      setView('student');
      // Pre-fetch geofence parameters for instant location checks
      fetch(`/api/sessions/lookup/${sessionParam}`)
        .then(res => res.json())
        .then(data => {
          if (data.found && data.geofence) {
            setClassGeofence(data.geofence);
          }
        })
        .catch(console.error);
      requestStudentLocation();
    }

    // Attempt to guess local network IP for Teacher QR code url
    const currentHost = window.location.hostname;
    if (currentHost === 'localhost' || currentHost === '127.0.0.1') {
      setTeacherCustomIp('');
    } else {
      setTeacherCustomIp(currentHost);
    }
  }, []);

  // Automatically recalculate location bounds when class geofence loads
  useEffect(() => {
    if (view === 'student' && classGeofence) {
      requestStudentLocation();
    }
  }, [classGeofence, view]);

  // Sync active classrooms list when setup screen is loaded
  useEffect(() => {
    if (view === 'teacher_setup') {
      fetch('/api/database/classrooms')
        .then(res => res.json())
        .then(data => {
          if (Array.isArray(data)) {
            setClassroomsList(data);
          }
        })
        .catch(console.error);
    }
  }, [view]);

  // Restore teacher active session from localStorage on page refresh/reload
  useEffect(() => {
    const savedSessionId = localStorage.getItem('sairam_active_session_id');
    const savedTeacherName = localStorage.getItem('sairam_active_teacher_name');
    const savedTeacherDept = localStorage.getItem('sairam_active_teacher_dept');
    const savedFacultyUser = localStorage.getItem('sairam_active_faculty_user');
    const savedSheetUrl = localStorage.getItem('sairam_active_sheet_url');

    if (savedSessionId && savedTeacherName && savedTeacherDept && savedFacultyUser) {
      // Confirm session is still active on the server
      fetch(`/api/sessions/${savedSessionId}/otp-status`)
        .then(res => {
          if (res.ok) return res.json();
          throw new Error("Session not found on server");
        })
        .then(() => {
          setFacultyUser(savedFacultyUser);
          setActiveTeacherName(savedTeacherName);
          setDepartment(savedTeacherDept);
          if (savedSheetUrl) setGoogleSheetUrl(savedSheetUrl);
          setClassCode(savedSessionId);
          
          // Re-fetch current students list & details securely
          fetch(`/api/sessions/${savedSessionId}?teacherUsername=${savedFacultyUser}`)
            .then(r => r.json())
            .then(sessionData => {
              setActiveSession(sessionData);
              setView('teacher_dashboard');
            });
        })
        .catch(() => {
          // Clear stale references if expired on server
          localStorage.removeItem('sairam_active_session_id');
        });
    }
  }, []);

  // Poll Teacher Dashboard details in real-time
  useEffect(() => {
    if (view === 'teacher_dashboard' && activeSession) {
      pollingRef.current = window.setInterval(() => {
        fetchSessionDetails(activeSession.id);
      }, 3000);

      return () => {
        if (pollingRef.current) {
          clearInterval(pollingRef.current);
        }
      };
    }
  }, [view, activeSession?.id]);

  // Poll Student Approval Status if waiting for device sharing approval
  useEffect(() => {
    if (view === 'student' && pendingApproval && classCode) {
      approvalPollingRef.current = window.setInterval(() => {
        checkApprovalStatus();
      }, 3000);

      return () => {
        if (approvalPollingRef.current) {
          clearInterval(approvalPollingRef.current);
        }
      };
    }
  }, [view, pendingApproval, classCode, studentRegNo]);

  // OTP Countdown/Refresh Ring Animation for TV Screen
  useEffect(() => {
    if (view === 'teacher_dashboard' && activeSession) {
      const interval = setInterval(() => {
        setOtpProgress((prev) => {
          if (prev <= 10) {
            return 100; // Reset
          }
          return prev - 1; // Decrement
        });
      }, 100);

      return () => clearInterval(interval);
    }
  }, [view, activeSession]);

  // If OTP updates on fetch, reset progress ring
  useEffect(() => {
    setOtpProgress(100);
  }, [activeSession?.currentOtp]);

  // Lookup student dynamically from Student ID input (Offline-first caching)
  useEffect(() => {
    if (studentRegNo.length >= 5) {
      // Check local offline cache first for instant loading
      const cached = localStorage.getItem('sairam_student_cached_profile');
      if (cached) {
        try {
          const profile = JSON.parse(cached);
          if (profile.regNo.toLowerCase().trim() === studentRegNo.toLowerCase().trim()) {
            setStudentName(profile.name);
            setStudentMatchingMsg(`✅ Directory Match (Cached): ${profile.name} (${profile.dept} Sec ${profile.sec})`);
            return;
          }
        } catch (e) {
          console.error("Local profile cache parse failed", e);
        }
      }

      const delayDebounce = setTimeout(() => {
        fetch(`/api/students/search?q=${studentRegNo}`)
          .then(res => res.json())
          .then((data: any[]) => {
            if (data.length > 0) {
              const matched = data.find(s => s.regNo.toLowerCase() === studentRegNo.toLowerCase()) || data[0];
              setStudentName(matched.name);
              setStudentMatchingMsg(`✅ Directory Match: ${matched.name} (${matched.dept} Sec ${matched.sec})`);
            } else {
              setStudentMatchingMsg('❌ Student ID not found in college database.');
              setStudentName('');
            }
          })
          .catch(() => {
            setStudentMatchingMsg('');
          });
      }, 300);

      return () => clearTimeout(delayDebounce);
    } else {
      setStudentMatchingMsg('');
      setStudentName('');
    }
  }, [studentRegNo]);

  // Real-time Class Access Code validation lookup (Debounced Room check)
  useEffect(() => {
    if (classCode.length >= 3) {
      setIsClassSearching(true);
      setClassLookupError('');
      const delayDebounce = setTimeout(() => {
        fetch(`/api/sessions/lookup/${classCode}`)
          .then(res => res.json())
          .then(data => {
            setIsClassSearching(false);
            if (data.found) {
              setClassLookup({
                classroomCode: data.classroomCode,
                department: data.department,
                section: data.section,
                year: data.year,
                teacherName: data.teacherName
              });
              if (data.geofence) {
                setClassGeofence(data.geofence);
              }
              // Automatically request location when class is verified
              requestStudentLocation();
            } else {
              setClassLookup(null);
              setClassLookupError('❌ Classroom session not active. Check room code.');
            }
          })
          .catch(() => {
            setIsClassSearching(false);
            setClassLookup(null);
            setClassLookupError('⚠️ Network connection failed.');
          });
      }, 400);

      return () => clearTimeout(delayDebounce);
    } else {
      setClassLookup(null);
      setClassLookupError('');
    }
  }, [classCode]);

  const startScanner = async () => {
    setShowScanner(true);
    setScannerError('');
    
    // Wait for the modal element #qr-reader to mount in DOM
    setTimeout(async () => {
      try {
        const html5Qrcode = new Html5Qrcode("qr-reader");
        scannerRef.current = html5Qrcode;

        await html5Qrcode.start(
          { facingMode: "environment" },
          {
            fps: 10,
            qrbox: { width: 250, height: 250 }
          },
          (decodedText) => {
            handleQrSuccess(decodedText);
          },
          () => {
            // silent fail for constant polling
          }
        );

        // Fetch zoom capabilities once the stream is running
        setTimeout(() => {
          const track = getCameraTrack();
          if (track) {
            try {
              const capabilities = track.getCapabilities() as any;
              if (capabilities && capabilities.zoom) {
                setZoomSupported(true);
                setMinZoom(capabilities.zoom.min || 1);
                setMaxZoom(capabilities.zoom.max || 4);
                setCurrentZoom(capabilities.zoom.min || 1);
              } else {
                setZoomSupported(false);
              }
            } catch (err) {
              console.warn("Could not read camera capabilities", err);
              setZoomSupported(false);
            }
          }
        }, 1200);
      } catch (err: any) {
        console.error("Camera access error", err);
        setScannerError("Camera access denied. Please grant camera permission.");
      }
    }, 200);
  };

  const stopScanner = async () => {
    setZoomSupported(false);
    if (scannerRef.current) {
      if (scannerRef.current.isScanning) {
        try {
          await scannerRef.current.stop();
        } catch (e) {
          console.error("Error stopping scanner", e);
        }
      }
      scannerRef.current = null;
    }
    setShowScanner(false);
  };

  const handleQrSuccess = (decodedText: string) => {
    try {
      const url = new URL(decodedText);
      const sessionParam = url.searchParams.get('session');
      const tokenParam = url.searchParams.get('token');

      if (sessionParam) {
        setClassCode(sessionParam);
        if (tokenParam) {
          setScannedToken(tokenParam);
          setStudentOtp('QR-SCAN');
        }
        stopScanner();
      } else {
        setScannerError("Invalid QR Code: Session ID missing.");
      }
    } catch (e) {
      setScannerError("Invalid QR Code link format.");
    }
  };

  const handleFacultyLogin = async (e: React.FormEvent) => {
    e.preventDefault();
    setFacultyLoginError('');
    setIsLoggingIn(true);
    try {
      let res = await fetch('/api/faculty/login', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ username: facultyUser.trim(), password: facultyPass.trim() })
      });
      let data = await res.json();
      
      if (!res.ok) {
        // If login failed, check if the server database was wiped due to serverless cold start
        const statusRes = await fetch('/api/database/status');
        const statusData = await statusRes.json();
        
        if (statusData.empty || statusData.facultyCount <= 1) {
          const cachedRosterStr = localStorage.getItem('sairam_roster_data');
          if (cachedRosterStr) {
            const cachedRoster = JSON.parse(cachedRosterStr);
            console.log("[Login Self-Healing] Roster cache missing. Restoring database...");
            const restoreRes = await fetch('/api/database/restore', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({
                faculty: cachedRoster.faculty || [],
                students: cachedRoster.students || [],
                classrooms: cachedRoster.classrooms || []
              })
            });
            
            if (restoreRes.ok) {
              console.log("[Login Self-Healing] Database restored successfully. Retrying login...");
              res = await fetch('/api/faculty/login', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ username: facultyUser.trim(), password: facultyPass.trim() })
              });
              data = await res.json();
              if (!res.ok) throw new Error(data.error || 'Authentication failed');
            } else {
              throw new Error(data.error || 'Authentication failed');
            }
          } else {
            throw new Error(data.error || 'Authentication failed');
          }
        } else {
          throw new Error(data.error || 'Authentication failed');
        }
      }
      
      setActiveTeacherName(data.name);
      setDepartment(data.department === 'ALL' ? 'IT' : data.department);
      setView('teacher_setup');
    } catch (err: any) {
      setFacultyLoginError(err.message);
    } finally {
      setIsLoggingIn(false);
    }
  };

  // Cleanup scanner stream and GPS watch on component unmount
  useEffect(() => {
    return () => {
      if (scannerRef.current && scannerRef.current.isScanning) {
        scannerRef.current.stop().catch(console.error);
      }
      if (gpsWatchRef.current !== null) {
        navigator.geolocation.clearWatch(gpsWatchRef.current);
      }
    };
  }, []);

  // Self-Healing Roster Recovery: Automatically restore database if server memory was wiped by Vercel reboot
  useEffect(() => {
    const restoreDatabase = async () => {
      try {
        const statusRes = await fetch('/api/database/status');
        const statusData = await statusRes.json();
        
        if (statusData.empty) {
          const cachedRosterStr = localStorage.getItem('sairam_roster_data');
          if (cachedRosterStr) {
            const cachedRoster = JSON.parse(cachedRosterStr);
            console.log("[Self-Healing] Roster cache missing on server. Restoring database...");
            const restoreRes = await fetch('/api/database/restore', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({
                faculty: cachedRoster.faculty || [],
                students: cachedRoster.students || [],
                classrooms: cachedRoster.classrooms || []
              })
            });
            if (restoreRes.ok) {
              console.log("[Self-Healing] Database restored successfully!");
              // Refresh classrooms list locally since it was updated
              fetch('/api/database/classrooms')
                .then(r => r.json())
                .then(list => { if (Array.isArray(list)) setClassroomsList(list); })
                .catch(() => {});
            }
          }
        }
      } catch (e) {
        console.warn("[Self-Healing] Error recovering database status:", e);
      }
    };
    restoreDatabase();
  }, []);

  // ==========================================
  // API CALLS & ACTIONS
  // ==========================================

  // Create a new session (Teacher setup)
  const handleCreateSession = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!teacherClassCode) {
      alert("Please enter or select a Classroom Code");
      return;
    }
    try {
      const res = await fetch('/api/sessions', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          creatorUsername: facultyUser,
          classCode: teacherClassCode.trim().toUpperCase(),
          department,
          section,
          year,
          teacherName: activeTeacherName,
          sessionType,
          geofence: {
            latitude: parseFloat(geofenceLat) || SAIRAM_CAMPUS_COORDS.latitude,
            longitude: parseFloat(geofenceLng) || SAIRAM_CAMPUS_COORDS.longitude,
            radius: parseFloat(geofenceRadius) || 30
          },
          googleSheetUrl
        })
      });

      if (!res.ok) {
        const errorData = await res.json();
        throw new Error(errorData.error || "Failed to create session");
      }
      const data = await res.json();
      setActiveSession(data);
      
      // Save details to LocalStorage to protect session against reloads/closures
      localStorage.setItem('sairam_active_session_id', data.id);
      localStorage.setItem('sairam_active_teacher_name', activeTeacherName);
      localStorage.setItem('sairam_active_teacher_dept', department);
      localStorage.setItem('sairam_active_faculty_user', facultyUser);
      localStorage.setItem('sairam_active_sheet_url', googleSheetUrl);
      
      setView('teacher_dashboard');
    } catch (err: any) {
      alert("Error starting class session: " + err.message);
    }
  };

  // Sync Faculty and Student roster database from Google Sheets Web App URL
  const handleSyncRoster = async () => {
    if (!rosterUrl) {
      setRosterSyncMsg("❌ Please enter a valid Web App URL");
      return;
    }
    
    setIsSyncingRoster(true);
    setRosterSyncMsg("");
    localStorage.setItem('sairam_roster_url', rosterUrl);
    
    try {
      const res = await fetch('/api/database/sync-roster', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ rosterDbUrl: rosterUrl })
      });
      const data = await res.json();
      
      if (res.ok) {
        setRosterSyncMsg(`✓ Loaded ${data.facultyCount} teachers, ${data.studentCount} students & ${data.classroomCount || 0} classrooms!`);
        
        // Save database cache to localStorage for automatic Self-Healing recovery
        localStorage.setItem('sairam_roster_data', JSON.stringify({
          faculty: data.faculty || [],
          students: data.students || [],
          classrooms: data.classrooms || []
        }));

        // Proactively refresh classrooms list locally
        fetch('/api/database/classrooms')
          .then(r => r.json())
          .then(list => {
            if (Array.isArray(list)) setClassroomsList(list);
          })
          .catch(() => {});
      } else {
        setRosterSyncMsg(`❌ Sync failed: ${data.error || 'Server error'}`);
      }
    } catch (e: any) {
      setRosterSyncMsg(`❌ Connection error: ${e.message}`);
    } finally {
      setIsSyncingRoster(false);
    }
  };

  // Handle classroom selector dropdown selection change
  const handleClassroomSelection = (indexStr: string) => {
    setSelectedClassroomIndex(indexStr);
    if (indexStr !== '') {
      const index = parseInt(indexStr);
      const room = classroomsList[index];
      if (room) {
        setTeacherClassCode(room.classCode);
        setGeofenceLat(room.latitude.toString());
        setGeofenceLng(room.longitude.toString());
        setGeofenceRadius((room.radius || "30").toString());
      }
    } else {
      setTeacherClassCode('');
    }
  };


  // Submit faculty password change request to backend (with robust JSON response validation bypassing CORS content-type masking)
  const handlePasswordChange = async (e: React.FormEvent) => {
    e.preventDefault();
    if (newPassword !== confirmNewPassword) {
      setChangePasswordMsg("❌ New passwords do not match!");
      return;
    }
    
    setIsChangingPassword(true);
    setChangePasswordMsg("");
    
    try {
      const res = await fetch('/api/faculty/change-password', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          username: facultyUser || 'admin',
          currentPassword,
          newPassword
        })
      });
      
      let data;
      const text = await res.text();
      try {
        data = JSON.parse(text);
      } catch (err) {
        setChangePasswordMsg("❌ Server Error: Received non-JSON response from server.");
        console.error("Non-JSON response text:", text);
        return;
      }

      if (res.ok) {
        setChangePasswordMsg("✓ Password updated successfully!");
        setCurrentPassword('');
        setNewPassword('');
        setConfirmNewPassword('');
        setTimeout(() => {
          setShowChangePasswordModal(false);
          setChangePasswordMsg('');
        }, 1500);
      } else {
        setChangePasswordMsg(`❌ Error: ${data.error || 'Failed to update password'}`);
      }
    } catch (e: any) {
      setChangePasswordMsg(`❌ Connection error: ${e.message}`);
    } finally {
      setIsChangingPassword(false);
    }
  };

  // Automatically detect teacher's current GPS location
  const detectTeacherLocation = () => {
    if (!navigator.geolocation) {
      alert("Geolocation is not supported by your browser");
      return;
    }
    navigator.geolocation.getCurrentPosition(
      (position) => {
        setGeofenceLat(position.coords.latitude.toFixed(6));
        setGeofenceLng(position.coords.longitude.toFixed(6));
      },
      (error) => {
        alert(`Failed to detect location: ${error.message}. Please verify GPS permissions.`);
      },
      { enableHighAccuracy: true, timeout: 8000 }
    );
  };

  // Open edit student dialog
  const openEditStudent = (st: StudentAttendance) => {
    setEditingStudent(st);
    setEditStudentName(st.name);
    setEditStudentReg(st.regNo);
    setShowEditStudentModal(true);
  };

  // Submit student edits to backend
  const handleSaveEditStudent = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!activeSession || !editingStudent) return;
    try {
      const res = await fetch(`/api/sessions/${activeSession.id}/edit-student`, {
        method: 'POST',
        headers: { 
          'Content-Type': 'application/json',
          'X-Teacher-Username': facultyUser
        },
        body: JSON.stringify({
          oldRegNo: editingStudent.regNo,
          regNo: editStudentReg.trim(),
          name: editStudentName.trim()
        })
      });
      const data = await res.json();
      if (res.ok) {
        fetchSessionDetails(activeSession.id);
        setShowEditStudentModal(false);
        setEditingStudent(null);
      } else {
        alert(data.error || "Failed to update student details");
      }
    } catch (err) {
      alert("Network error: Failed to update student details");
    }
  };

  // Submit toggled session type to backend
  const handleToggleSessionType = async (type: 'IN' | 'OUT') => {
    if (!activeSession) return;
    try {
      const res = await fetch(`/api/sessions/${activeSession.id}/session-type`, {
        method: 'PUT',
        headers: { 
          'Content-Type': 'application/json',
          'X-Teacher-Username': facultyUser
        },
        body: JSON.stringify({ sessionType: type })
      });
      if (res.ok) {
        // Refresh session details to load the updated type and OTP
        fetchSessionDetails(activeSession.id);
      }
    } catch (e) {
      console.error("Failed to toggle session type:", e);
    }
  };

  // Fetch session details for dashboard update (Authenticated with creator Isolation credentials)
  const fetchSessionDetails = async (id: string) => {
    try {
      const res = await fetch(`/api/sessions/${id}?teacherUsername=${facultyUser}`);
      if (res.ok) {
        const data = await res.json();
        setActiveSession(data);
      }
    } catch (err) {
      console.error("Failed to poll session details", err);
    }
  };
  // Student location verification request (with watchPosition settling coordinate filters)
  const requestStudentLocation = () => {
    setGeoStatus('checking');
    if (!navigator.geolocation) {
      setGeoStatus('denied');
      setAttendanceError("Geolocation is not supported by your browser");
      return;
    }

    if (gpsWatchRef.current !== null) {
      navigator.geolocation.clearWatch(gpsWatchRef.current);
      gpsWatchRef.current = null;
    }

    let bestCoords: { latitude: number; longitude: number } | null = null;
    let timerId: number | null = null;

    const stopWatching = () => {
      if (gpsWatchRef.current !== null) {
        navigator.geolocation.clearWatch(gpsWatchRef.current);
        gpsWatchRef.current = null;
      }
      if (timerId !== null) {
        window.clearTimeout(timerId);
        timerId = null;
      }
    };

    const startWatching = (highAccuracy: boolean) => {
      gpsWatchRef.current = navigator.geolocation.watchPosition(
        (position) => {
          const { latitude, longitude, accuracy } = position.coords;
          setStudentCoords({ lat: latitude, lng: longitude });
          bestCoords = { latitude, longitude };
          console.log(`GPS stream reading (${highAccuracy ? 'High' : 'Low'} Accuracy): Accuracy ${accuracy}m at ${latitude}, ${longitude}`);

          const targetLat = classGeofence?.latitude || SAIRAM_CAMPUS_COORDS.latitude;
          const targetLng = classGeofence?.longitude || SAIRAM_CAMPUS_COORDS.longitude;
          const targetRadius = classGeofence?.radius || 500;
          
          const check = verifyGeofence(latitude, longitude, targetLat, targetLng, targetRadius);
          setStudentDistance(check.distance);

          if (check.inGeofence) {
            setGeoStatus('success');
            setAttendanceError('');
            stopWatching();
          } else {
            setGeoStatus('out_of_bounds');
            setAttendanceError(`Geofence Error: You are outside classroom boundary. Distance: ${Math.round(check.distance)}m (Limit: ${targetRadius}m)`);
          }
        },
        (error) => {
          console.error(`GPS Watch stream error (HighAccuracy: ${highAccuracy})`, error);
          if (highAccuracy) {
            console.log("High accuracy timed out or unavailable. Switching to standard accuracy...");
            stopWatching();
            startWatching(false);
          } else {
            if (!bestCoords) {
              setGeoStatus('denied');
              setAttendanceError(`GPS Location Error: ${error.message}. Please check location permissions.`);
            }
          }
        },
        { enableHighAccuracy: highAccuracy, timeout: 6000, maximumAge: 0 }
      );
    };

    // Timeout of 15 seconds to allow mobile GPS hardware to warm up and settle
    timerId = window.setTimeout(() => {
      if (!bestCoords) {
        stopWatching();
        setGeoStatus('denied');
        setAttendanceError("GPS Location timeout. Unable to secure a clear satellite lock. Please ensure location services are enabled.");
      } else {
        const targetLat = classGeofence?.latitude || SAIRAM_CAMPUS_COORDS.latitude;
        const targetLng = classGeofence?.longitude || SAIRAM_CAMPUS_COORDS.longitude;
        const targetRadius = classGeofence?.radius || 500;
        const check = verifyGeofence(bestCoords.latitude, bestCoords.longitude, targetLat, targetLng, targetRadius);
        
        setStudentDistance(check.distance);
        if (check.inGeofence) {
          setGeoStatus('success');
          setAttendanceError('');
          stopWatching();
        } else {
          setGeoStatus('out_of_bounds');
          setAttendanceError(`Geofence Error: You are outside classroom boundary. Distance: ${Math.round(check.distance)}m (Limit: ${targetRadius}m)`);
          stopWatching();
        }
      }
    }, 15000);

    // Start with high accuracy initially
    startWatching(true);
  };

  // Student attendance submit
  const handleStudentSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (classCode.length !== 5) {
      setAttendanceError("Please enter a valid 5-digit Class Access Code");
      return;
    }
    if (!studentMatchingMsg.includes('Directory Match')) {
      setAttendanceError("Student ID not found. Registration is closed to non-roster students.");
      return;
    }
    if (geoStatus !== 'success' && geoStatus !== 'out_of_bounds') {
      setAttendanceError("GPS Geofence status must be verified first");
      return;
    }
    if (geoStatus === 'out_of_bounds') {
      setAttendanceError(`Geofencing block. You are outside the Sairam College campus. (Distance: ${studentDistance ? Math.round(studentDistance) : 'Unknown'}m)`);
      return;
    }

    setAttendanceError('');
    setIsSubmitting(true);

    try {
      let res: Response | null = null;
      let retries = 3;
      let success = false;
      let lastError = null;

      while (retries > 0 && !success) {
        try {
          res = await fetch(`/api/sessions/${classCode}/attend`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              regNo: studentRegNo.trim(),
              name: studentName.trim(),
              rollNo: "",
              otp: studentOtp === 'QR-SCAN' ? undefined : studentOtp.trim(),
              token: scannedToken || undefined,
              location: studentCoords,
              fingerprint: studentDeviceFingerprint
            })
          });
          success = true;
        } catch (err: any) {
          retries--;
          lastError = err;
          if (retries > 0) {
            // Wait 1.5s before retrying
            await new Promise(resolve => setTimeout(resolve, 1500));
          }
        }
      }

      if (!res || !success) {
        throw lastError || new Error("Failed to reach server after retries.");
      }

      const data = await res.json();
      setIsSubmitting(false);

      if (res.status === 202 && data.status === 'pending_approval') {
        setPendingApproval(true);
        setAttendanceError('');
      } else if (!res.ok) {
        setAttendanceError(data.error || "Failed to mark attendance");
      } else if (data.status === 'success') {
        // Cache successful profile locally for instant offline loading next time
        localStorage.setItem('sairam_student_cached_profile', JSON.stringify({
          regNo: studentRegNo.trim(),
          rollNo: "",
          name: studentName.trim(),
          dept: classLookup?.department || "",
          sec: classLookup?.section || "",
          year: classLookup?.year || ""
        }));
        setMarkedStudentDetails(data.student);
        setView('student_success');
      }
    } catch (err: any) {
      setIsSubmitting(false);
      setAttendanceError("Connection to server failed. Please check network.");
    }
  };

  // Student polls approval status securely (Only checks personal approval status, avoiding full details leak)
  const checkApprovalStatus = async () => {
    if (!classCode || !studentRegNo) return;
    try {
      const res = await fetch(`/api/sessions/${classCode}/approval-status/${studentRegNo}`);
      if (res.ok) {
        const data = await res.json();
        
        if (data.approved) {
          setMarkedStudentDetails({
            regNo: studentRegNo,
            rollNo: "",
            name: studentName,
            timestamp: new Date().toLocaleTimeString(),
            deviceStatus: "Shared (Approved)",
            method: "Teacher Approved"
          });
          setPendingApproval(false);
          if (approvalPollingRef.current) clearInterval(approvalPollingRef.current);
          setView('student_success');
        } else if (!data.pending) {
          setPendingApproval(false);
          if (approvalPollingRef.current) clearInterval(approvalPollingRef.current);
          setAttendanceError("Device sharing request rejected by Teacher.");
        }
      }
    } catch (e) {
      console.error(e);
    }
  };

  // Teacher manual student add
  const handleManualMark = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!activeSession || !searchRegNo) return;

    try {
      const res = await fetch(`/api/sessions/${activeSession.id}/manual-mark`, {
        method: 'POST',
        headers: { 
          'Content-Type': 'application/json',
          'X-Teacher-Username': facultyUser
        },
        body: JSON.stringify({ regNo: searchRegNo.trim() })
      });

      const data = await res.json();
      if (!res.ok) {
        alert(data.error || "Failed to mark student manually");
      } else {
        fetchSessionDetails(activeSession.id);
        setSearchRegNo('');
        alert(data.message);
      }
    } catch (err) {
      alert("Error making request");
    }
  };

  // Teacher approves device sharing
  const handleApproveSharing = async (regNo: string) => {
    if (!activeSession) return;
    try {
      const res = await fetch(`/api/sessions/${activeSession.id}/approve-sharing`, {
        method: 'POST',
        headers: { 
          'Content-Type': 'application/json',
          'X-Teacher-Username': facultyUser
        },
        body: JSON.stringify({ regNo })
      });
      if (res.ok) {
        fetchSessionDetails(activeSession.id);
      } else {
        const err = await res.json();
        alert(err.error);
      }
    } catch (e) {
      alert("Error approving sharing request");
    }
  };

  // Teacher rejects device sharing
  const handleRejectSharing = async (regNo: string) => {
    if (!activeSession) return;
    try {
      const res = await fetch(`/api/sessions/${activeSession.id}/reject-sharing`, {
        method: 'POST',
        headers: { 
          'Content-Type': 'application/json',
          'X-Teacher-Username': facultyUser
        },
        body: JSON.stringify({ regNo })
      });
      if (res.ok) {
        fetchSessionDetails(activeSession.id);
      }
    } catch (e) {
      alert("Error rejecting sharing request");
    }
  };

  // Sync to Google Sheets
  const handleSyncToSheets = async () => {
    if (!activeSession) return;
    setIsSyncing(true);
    setSyncStatus(null);

    try {
      const res = await fetch(`/api/sessions/${activeSession.id}/sync-sheets`, {
        method: 'POST',
        headers: { 
          'Content-Type': 'application/json',
          'X-Teacher-Username': facultyUser
        },
        body: JSON.stringify({ googleSheetUrl })
      });

      const data = await res.json();
      setIsSyncing(false);

      if (res.ok) {
        setSyncStatus({ success: true, message: data.message });
      } else {
        setSyncStatus({ success: false, message: data.error || "Failed to sync" });
      }
    } catch (e: any) {
      setIsSyncing(false);
      setSyncStatus({ success: false, message: `Sheet server connection failed: ${e.message}` });
    }
  };

  // Close Session
  const handleCloseSession = async () => {
    if (!activeSession) return;
    if (!window.confirm("Are you sure you want to end this attendance marking session? All current data will be cleared from server memory.")) return;

    try {
      await fetch(`/api/sessions/${activeSession.id}?teacherUsername=${facultyUser}`, { method: 'DELETE' });
      setActiveSession(null);
      localStorage.removeItem('sairam_active_session_id');
      setView('landing');
    } catch (e) {
      setView('landing');
    }
  };

  // Helper: Generates QR scanned url for student browser access
  const getQrUrl = () => {
    if (!activeSession) return '';
    const baseHost = teacherCustomIp || window.location.hostname;
    const basePort = window.location.port ? `:${window.location.port}` : '';
    return `${window.location.protocol}//${baseHost}${basePort}/?session=${activeSession.id}&token=${activeSession.currentToken}`;
  };

  // ==========================================
  // VIEW RENDER LOGIC
  // ==========================================

  return (
    <div className="min-h-screen flex flex-col justify-between">
      {/* Header */}
      <header className="glass-panel border-t-0 border-x-0 rounded-t-none rounded-b-xl py-4 px-6 flex items-center justify-between mx-auto w-full max-w-7xl mt-0">
        <div className="flex items-center gap-3">
          <img 
            src="https://blogger.googleusercontent.com/img/b/R29vZ2xl/AVvXsEiuC2k9ut9RfGb85tucvxFCE9go2EfdjrNJ7dtxfFP9jQxzm-D0vhLqekJPOut7Fky2empX1vyFGX94MPw7ZtdBywU5Q9fPWhaO8lDsIAuoo5FDFGgMMMlnLwPIXU-6UCpTs1CJlaf9XSPHdXd6rMCaMsROAkDz6_CNuMIwHXmDNO01xAQgeoYbkEACARE/s191/Sri%20Sairam%20Engineering%20College,%20Chennai.png" 
            alt="Sri Sairam Engineering College Logo" 
            className="w-10 h-10 object-contain bg-white rounded-lg p-0.5 shadow-md flex-shrink-0"
          />
          <div className="text-left">
            <h1 className="text-base md:text-lg font-extrabold tracking-tight text-text-primary">
              SRI SAIRAM ENGINEERING COLLEGE
            </h1>
            <p className="text-xs text-text-muted font-bold tracking-wide uppercase">ERP Attendance Hub</p>
          </div>
        </div>
        
        <div className="flex items-center gap-2">
          {(view === 'teacher_setup' || view === 'teacher_dashboard') && (
            <button
              onClick={() => {
                setChangePasswordMsg('');
                setCurrentPassword('');
                setNewPassword('');
                setConfirmNewPassword('');
                setShowChangePasswordModal(true);
              }}
              className="btn-secondary text-[11px] py-1.5 px-3 rounded-lg flex items-center gap-1.5 font-bold shadow-sm"
            >
              <Lock size={12} /> Change Password
            </button>
          )}
          {view !== 'landing' && (
            <button 
              onClick={() => {
                if (view === 'teacher_dashboard') {
                  handleCloseSession();
                } else {
                  setView('landing');
                  setClassCode('');
                  setStudentRegNo('');
                  setStudentOtp('');
                  setScannedToken('');
                  setGeoStatus('idle');
                  setAttendanceError('');
                  setPendingApproval(false);
                }
              }}
              className="btn-secondary text-[11px] py-1.5 px-3 rounded-lg flex items-center gap-1.5 font-bold"
            >
              Leave Portal
            </button>
          )}
        </div>
      </header>

      {/* Main Content Area */}
      <main className="flex-grow flex items-center justify-center p-4 max-w-7xl mx-auto w-full">
        
        {/* ==========================================
            VIEW 1: LANDING PAGE (ROLE CHOICE)
            ========================================== */}
        {view === 'landing' && (
          <div className="text-center w-full max-w-4xl py-8 md:py-16 slide-up">
            <span className="text-xs font-bold tracking-widest badge-cyan px-3.5 py-1.5 rounded-full uppercase glow-text-cyan">
              Sairam ERP Core
            </span>
            <h2 className="text-3xl md:text-5xl font-black mt-6 tracking-tight text-text-primary">
              Institutional Geofenced <br />
              <span className="bg-gradient-to-r from-primary to-secondary bg-clip-text text-transparent">
                Attendance Marking System
              </span>
            </h2>
            <p className="text-text-secondary text-sm md:text-base max-w-lg mx-auto mt-4 leading-relaxed">
              Verify student check-ins securely using time-synchronized rotating OTP codes, browser GPS positioning, and device-bound anti-cheat registers.
            </p>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-6 mt-12 max-w-2xl mx-auto">
              {/* Teacher Portal Trigger */}
              <div 
                onClick={() => {
                  setFacultyUser('');
                  setFacultyPass('');
                  setFacultyLoginError('');
                  setView('teacher_login');
                }}
                className="glass-panel p-8 text-left cursor-pointer group flex flex-col justify-between h-64 border border-white/5 hover:border-cyan-500/40 relative overflow-hidden"
              >
                <div className="absolute right-0 top-0 w-24 h-24 bg-primary/5 rounded-full filter blur-xl group-hover:bg-primary/10 transition-all duration-300"></div>
                <div className="w-12 h-12 rounded-xl bg-primary/10 border border-primary/20 flex items-center justify-center text-cyan-400 group-hover:scale-110 transition-transform duration-300">
                  <Tv size={24} className="text-cyan-400" />
                </div>
                <div>
                  <h3 className="text-xl font-bold text-text-primary group-hover:text-cyan-400 transition-colors">
                    Teacher Portal
                  </h3>
                  <p className="text-xs text-text-secondary mt-2">
                    Create geofenced classes, broadcast dynamic QR codes on classroom TV screens, and manage real-time present lists and device sharing queues.
                  </p>
                </div>
                <div className="flex items-center gap-2 text-xs font-bold text-cyan-400 mt-4">
                  Open Control Board <ChevronRight size={14} />
                </div>
              </div>

              {/* Student Portal Trigger */}
              <div 
                onClick={() => {
                  setView('student');
                  requestStudentLocation();
                }}
                className="glass-panel p-8 text-left cursor-pointer group flex flex-col justify-between h-64 border border-white/5 hover:border-cyan-500/40 relative overflow-hidden"
              >
                <div className="absolute right-0 top-0 w-24 h-24 bg-secondary/5 rounded-full filter blur-xl group-hover:bg-secondary/10 transition-all duration-300"></div>
                <div className="w-12 h-12 rounded-xl bg-secondary/10 border border-secondary/20 flex items-center justify-center text-purple-400 group-hover:scale-110 transition-transform duration-300">
                  <QrCode size={24} className="text-yellow-400" />
                </div>
                <div>
                  <h3 className="text-xl font-bold text-text-primary group-hover:text-yellow-400 transition-colors">
                    Student Login
                  </h3>
                  <p className="text-xs text-text-secondary mt-2">
                    Enter the active classroom code or scan the Smart TV screen. Requires coordinates verification within Sairam campus grounds.
                  </p>
                </div>
                <div className="flex items-center gap-2 text-xs font-bold text-yellow-400 mt-4">
                  Mark Present Now <ChevronRight size={14} />
                </div>
              </div>
            </div>
          </div>
        )}

        {/* ==========================================
            VIEW 1.5: TEACHER FACULTY LOGIN
            ========================================== */}
        {view === 'teacher_login' && (
          <div className="w-full max-w-md py-8 slide-up">
            <div className="glass-panel p-8 border border-white/5 relative">
              <div className="border-b border-color pb-4 mb-6 text-center">
                <h3 className="text-2xl font-black text-text-primary">Faculty Authentication</h3>
                <p className="text-xs text-text-secondary mt-1">Access Sri Sairam Engineering College ERP</p>
              </div>

              {facultyLoginError && (
                <div className="p-3 bg-red-500/5 border border-red-500/15 text-red-400 text-xs rounded-lg mb-4 flex items-start gap-2">
                  <AlertTriangle size={14} className="flex-shrink-0 mt-0.5 text-red-400" />
                  <span>{facultyLoginError}</span>
                </div>
              )}

              <form onSubmit={handleFacultyLogin} className="space-y-4">
                <div>
                  <label className="block text-xs font-semibold text-text-secondary uppercase mb-1">Faculty Username / ID</label>
                  <input 
                    type="text" 
                    placeholder="e.g. admin" 
                    value={facultyUser} 
                    onChange={e => setFacultyUser(e.target.value)}
                    className="glass-input" 
                    required 
                  />
                </div>

                <div>
                  <label className="block text-xs font-semibold text-text-secondary uppercase mb-1">Password</label>
                  <input 
                    type="password" 
                    placeholder="Enter password" 
                    value={facultyPass} 
                    onChange={e => setFacultyPass(e.target.value)}
                    className="glass-input" 
                    required 
                  />
                </div>

                <button 
                  type="submit" 
                  disabled={isLoggingIn}
                  className="btn-primary w-full py-3.5 mt-4"
                >
                  {isLoggingIn ? (
                    <span className="flex items-center gap-2">
                      <Loader2 size={16} className="animate-spin text-white" /> Logging in...
                    </span>
                  ) : (
                    <span className="flex items-center gap-2 text-white">
                      <Lock size={16} className="text-white" /> Verify Credentials
                    </span>
                  )}
                </button>

                <button 
                  type="button" 
                  onClick={() => setView('landing')}
                  className="btn-secondary w-full py-3 text-xs mt-2"
                >
                  Back to Landing
                </button>
              </form>

              {/* Live Google Sheets Roster Sync Accordion */}
              <div className="mt-6 border-t border-color pt-4 text-left">
                <details className="group">
                  <summary className="text-[10px] font-bold text-text-secondary uppercase cursor-pointer flex justify-between items-center hover:text-text-primary select-none">
                    <span>🗄️ Connect College Roster Sheet (Custom DB)</span>
                    <span className="text-10 text-cyan-400 group-open:rotate-180 transition-transform">▼</span>
                  </summary>
                  <div className="mt-3 space-y-3">
                    {!isAdminUnlocked ? (
                      <div className="space-y-2 p-3 bg-slate-50 border border-color rounded-xl">
                        <label className="block text-[9px] font-bold text-text-secondary uppercase">Admin Password Required</label>
                        <div className="flex gap-2">
                          <input 
                            type="password"
                            placeholder="Enter Admin Password"
                            value={adminKey}
                            onChange={e => setAdminKey(e.target.value)}
                            className="glass-input py-1 text-xs"
                          />
                          <button 
                            type="button"
                            onClick={() => {
                              const expectedKey = import.meta.env.VITE_ADMIN_SYNC_PASSWORD || 'adminpassword';
                              if (adminKey === expectedKey) {
                                setIsAdminUnlocked(true);
                                setAdminKey('');
                              } else {
                                alert("Incorrect admin password!");
                              }
                            }}
                            className="btn-primary py-1 px-3.5 text-10 rounded-lg font-bold"
                          >
                            Unlock
                          </button>
                        </div>
                      </div>
                    ) : (
                      <>
                        <p className="text-[10px] text-text-muted leading-relaxed">
                          Deploy your custom roster sheet Web App to instantly sync faculty credentials and student roll numbers, replacing all mock datasets.
                        </p>
                        
                        <div className="space-y-2">
                          <label className="block text-[9px] font-semibold text-text-secondary uppercase">Roster Web App URL</label>
                          <input 
                            type="url"
                            placeholder="https://script.google.com/macros/s/..."
                            value={rosterUrl}
                            onChange={e => setRosterUrl(e.target.value)}
                            className="glass-input py-1.5 text-xs"
                          />
                          
                          <div className="flex gap-2">
                            <button 
                              type="button"
                              onClick={handleSyncRoster}
                              disabled={isSyncingRoster}
                              className="btn-secondary py-1.5 px-3 text-[10px] rounded-lg flex-1 font-bold"
                            >
                              {isSyncingRoster ? "Syncing..." : "Sync Database"}
                            </button>
                            <button
                              type="button"
                              onClick={() => setShowRosterGuideModal(true)}
                              className="btn-secondary py-1.5 px-3 text-[10px] rounded-lg font-bold"
                            >
                              Setup Guide
                            </button>
                          </div>
                          
                          {rosterSyncMsg && (
                            <div className={`p-2 rounded text-[10px] font-bold ${
                              rosterSyncMsg.startsWith('❌') || rosterSyncMsg.startsWith('⚠️')
                                ? 'bg-red-500/5 text-red-400 border border-red-500/10'
                                : 'bg-emerald-500/5 text-emerald-400 border border-emerald-500/10'
                            }`}>
                              {rosterSyncMsg}
                            </div>
                          )}
                        </div>
                      </>
                    )}
                  </div>
                </details>
              </div>

            </div>
          </div>
        )}

        {/* ==========================================
            VIEW 2: TEACHER SETUP FORM
            ========================================== */}
        {view === 'teacher_setup' && (
          <div className="w-full max-w-xl slide-up py-4">
            <div className="glass-panel p-8 border border-white/5 text-left">
              <div className="flex items-center gap-3 mb-6">
                <div className="w-10 h-10 rounded-lg bg-primary/10 border border-primary/20 flex items-center justify-center text-cyan-400">
                  <Settings size={20} className="text-cyan-400" />
                </div>
                <div>
                  <h3 className="text-2xl font-bold text-text-primary">Start New Session</h3>
                  <p className="text-xs text-text-secondary">Set class details, geofence, and spreadsheet sync settings</p>
                </div>
              </div>

              <form onSubmit={handleCreateSession} className="space-y-4">
                
                {/* Linked Classroom Coordinates Load Selector */}
                <div className="mb-2">
                  <label className="block text-xs font-semibold text-text-secondary uppercase mb-1">Link Classroom from Sheet Database</label>
                  <select
                    value={selectedClassroomIndex}
                    onChange={e => handleClassroomSelection(e.target.value)}
                    className="glass-input text-xs"
                  >
                    <option value="">-- Enter Class Code Manually --</option>
                    {classroomsList.map((room, idx) => (
                      <option key={room.classCode} value={idx}>
                        Room {room.classCode} ({room.radius}m Geofence)
                      </option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="block text-xs font-semibold text-text-secondary uppercase mb-1">Classroom Code (e.g. G3103)</label>
                  <input 
                    type="text" 
                    value={teacherClassCode} 
                    onChange={e => setTeacherClassCode(e.target.value.toUpperCase())}
                    placeholder="e.g. G3103"
                    className="glass-input uppercase" 
                    required 
                  />
                </div>

                <div className="grid grid-cols-3 gap-3">
                  <div>
                    <label className="block text-xs font-semibold text-text-secondary uppercase mb-1">Department</label>
                    <select 
                      value={department} 
                      onChange={e => setDepartment(e.target.value)} 
                      className="glass-input"
                    >
                      <option value="IT">IT</option>
                      <option value="CSE">CSE</option>
                      <option value="ECE">ECE</option>
                      <option value="EEE">EEE</option>
                      <option value="MECH">MECH</option>
                      <option value="CIVIL">CIVIL</option>
                    </select>
                  </div>
                  <div>
                    <label className="block text-xs font-semibold text-text-secondary uppercase mb-1">Year</label>
                    <select 
                      value={year} 
                      onChange={e => setYear(e.target.value)} 
                      className="glass-input"
                    >
                      <option value="1st">1st Year</option>
                      <option value="2nd">2nd Year</option>
                      <option value="3rd">3rd Year</option>
                      <option value="4th">4th Year</option>
                    </select>
                  </div>
                  <div>
                    <label className="block text-xs font-semibold text-text-secondary uppercase mb-1">Section</label>
                    <select 
                      value={section} 
                      onChange={e => setSection(e.target.value)} 
                      className="glass-input"
                    >
                      <option value="A">Sec A</option>
                      <option value="B">Sec B</option>
                      <option value="C">Sec C</option>
                      <option value="D">Sec D</option>
                    </select>
                  </div>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-3 mt-3">
                  <div>
                    <label className="block text-xs font-semibold text-text-secondary uppercase mb-1">Session Attendance Type</label>
                    <select
                      value={sessionType}
                      onChange={e => setSessionType(e.target.value as 'IN' | 'OUT')}
                      className="glass-input text-xs font-bold text-cyan-400"
                    >
                      <option value="IN">Entry (IN)</option>
                      <option value="OUT">Exit (OUT)</option>
                    </select>
                  </div>
                </div>

                <div>
                  <div className="flex justify-between items-center border-b border-white/5 pb-1 mt-4">
                    <h4 className="text-xs font-bold text-text-primary uppercase tracking-wide">Geofence Range Settings</h4>
                    <button
                      type="button"
                      onClick={detectTeacherLocation}
                      className="text-[10px] font-bold text-cyan-400 hover:underline flex items-center gap-1 border-none bg-transparent cursor-pointer"
                    >
                      <MapPin size={10} className="text-cyan-400" /> Detect My Location
                    </button>
                  </div>
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-3 mt-2">
                    <div>
                      <label className="block text-[10px] font-semibold text-text-secondary uppercase mb-1">Campus Latitude</label>
                      <input 
                        type="text" 
                        value={geofenceLat} 
                        onChange={e => setGeofenceLat(e.target.value)}
                        className="glass-input py-2 text-sm" 
                      />
                    </div>
                    <div>
                      <label className="block text-[10px] font-semibold text-text-secondary uppercase mb-1">Campus Longitude</label>
                      <input 
                        type="text" 
                        value={geofenceLng} 
                        onChange={e => setGeofenceLng(e.target.value)}
                        className="glass-input py-2 text-sm" 
                      />
                    </div>
                  </div>
                  <div className="mt-2">
                    <label className="block text-[10px] font-semibold text-text-secondary uppercase mb-1">Allowed Radius (Meters)</label>
                    <select 
                      value={geofenceRadius} 
                      onChange={e => setGeofenceRadius(e.target.value)} 
                      className="glass-input py-2 text-sm"
                    >
                      <option value="15">15m (Classroom Desk Bounds)</option>
                      <option value="30">30m (Room Block Bounds)</option>
                      <option value="150">150m (Strict Classroom Block)</option>
                      <option value="300">300m (Departmental Wings)</option>
                      <option value="500">500m (Main Sairam Campus)</option>
                      <option value="800">800m (Extended Campus & Grounds)</option>
                    </select>
                  </div>
                </div>

                <div>
                  <div className="flex justify-between items-center border-b border-white/5 pb-1 mt-4">
                    <h4 className="text-xs font-bold text-text-primary uppercase tracking-wide">Google Sheets API Webhook</h4>
                    <button 
                      type="button"
                      onClick={() => setShowAppsScriptModal(true)}
                      className="text-[10px] font-bold text-cyan-400 hover:underline flex items-center gap-1 border-none bg-transparent cursor-pointer"
                    >
                      <Clipboard size={10} /> View Setup Instructions
                    </button>
                  </div>
                  <div className="mt-2">
                    <label className="block text-[10px] font-semibold text-text-secondary uppercase mb-1">Google Apps Script Web App URL</label>
                    <input 
                      type="url" 
                      placeholder="https://script.google.com/macros/s/..." 
                      value={googleSheetUrl} 
                      onChange={e => setGoogleSheetUrl(e.target.value)}
                      className="glass-input py-2 text-sm" 
                    />
                    <p className="text-[10px] text-text-muted mt-1 leading-normal">
                      Connect to your Google Sheet to record verified students automatically.
                    </p>
                  </div>
                </div>

                <div className="pt-4 flex gap-3">
                  <button 
                    type="button" 
                    onClick={() => setView('landing')}
                    className="btn-secondary flex-1 py-3"
                  >
                    Cancel
                  </button>
                  <button 
                    type="submit" 
                    className="btn-primary flex-2 py-3"
                  >
                    <Tv size={18} /> Initialize Class Screen
                  </button>
                </div>
              </form>
            </div>
          </div>
        )}

        {/* ==========================================
            VIEW 3: TEACHER SMART TV DASHBOARD
            ========================================== */}
        {view === 'teacher_dashboard' && activeSession && (
          <div className="w-full grid grid-cols-1 lg:grid-cols-12 gap-6 py-2 slide-up">
            
            {/* Left Col: TV QR & OTP Screen (8 Cols) */}
            <div className="lg:col-span-8 flex flex-col gap-6">
              
              {/* Main TV Panel */}
              <div className="glass-panel p-8 text-center flex flex-col justify-between items-center relative overflow-hidden border-2 border-white/5 min-h-580 bg-white">
                
                {/* Accent Gold top line */}
                <div className="absolute top-0 left-0 right-0 h-1 bg-gradient-to-r from-primary via-blue-700 to-secondary"></div>
                
                {/* Header details */}
                <div className="w-full flex justify-between items-start border-b border-color pb-4">
                  <div className="text-left">
                    <span className="text-10 font-extrabold badge-cyan px-2.5 py-1 rounded-md uppercase tracking-wider">
                      LIVE CLASSROOM MARKING
                    </span>
                    <h3 className="text-2xl font-black mt-2 text-text-primary tracking-tight">Classroom: {activeSession.classroomCode}</h3>
                    <p className="text-xs text-text-secondary font-semibold">Faculty In-Charge: {activeSession.teacherName} • {activeSession.year} Year {activeSession.department} Sec {activeSession.section}</p>
                    
                    <div className="flex items-center gap-1.5 mt-2.5">
                      <span className="text-[10px] font-bold text-text-secondary uppercase mr-1">Marking Mode:</span>
                      <button
                        onClick={() => handleToggleSessionType('IN')}
                        className={`text-10 font-bold px-2.5 py-1 rounded-l-md border transition-all ${
                          activeSession.sessionType === 'IN' 
                            ? 'bg-cyan-500 text-white border-cyan-500 shadow-sm' 
                            : 'bg-transparent text-text-secondary border-color hover:bg-slate-50'
                        }`}
                      >
                        IN (Entry)
                      </button>
                      <button
                        onClick={() => handleToggleSessionType('OUT')}
                        className={`text-10 font-bold px-2.5 py-1 rounded-r-md border-y border-r transition-all ${
                          activeSession.sessionType === 'OUT' 
                            ? 'bg-amber-500 text-white border-amber-500 shadow-sm' 
                            : 'bg-transparent text-text-secondary border-color border-l-0 hover:bg-slate-50'
                        }`}
                      >
                        OUT (Exit)
                      </button>
                    </div>
                  </div>
                  
                  <div className="text-right">
                    <span className="text-10 text-text-muted font-bold block uppercase">Campus Geofence</span>
                    <span className="text-xs text-cyan-400 font-extrabold flex items-center gap-1.5 justify-end">
                      <MapPin size={12} className="text-cyan-400" /> {activeSession.geofence.radius}m Radius
                    </span>
                  </div>
                </div>

                {/* Core QR & OTP Area */}
                <div className="flex flex-col md:flex-row items-center gap-12 my-6">
                  {/* QR Code Container */}
                  <div className="relative">
                    <div className="absolute inset-0 bg-primary/5 blur-xl scale-110 opacity-70"></div>
                    <div className="bg-white p-6 rounded-3xl shadow-2xl relative border border-color glow-border-cyan pulsing-ring">
                      {getQrUrl() ? (
                        <QRCodeSVG 
                          value={getQrUrl()} 
                          size={450}
                          level="L"
                          fgColor="#0f172a"
                          includeMargin={true}
                          className="w-full max-w-[450px] h-auto"
                        />
                      ) : (
                        <div className="w-full max-w-[450px] aspect-square flex items-center justify-center text-text-primary">Generating QR...</div>
                      )}
                    </div>
                    
                    <div className="mt-3 bg-primary-glow border border-primary/10 rounded-lg py-1 px-3 inline-flex items-center gap-1.5 mx-auto">
                      <span className="text-10 font-bold text-text-secondary uppercase">Access Code:</span>
                      <span className="text-xs font-black text-cyan-400 font-mono tracking-wider">{activeSession.id}</span>
                    </div>
                  </div>

                  {/* OTP Block */}
                  <div className="flex flex-col items-center justify-center md:items-start text-center md:text-left">
                    <span className="text-xs font-extrabold badge-purple px-3 py-1.5 rounded-lg">
                      Time-Rotating OTP
                    </span>
                    
                    <h2 className="text-6xl md:text-7xl font-black tracking-widest text-text-primary mt-4 glow-text-purple font-mono">
                      {activeSession.currentOtp ? (
                        `${activeSession.currentOtp.slice(0, 3)} ${activeSession.currentOtp.slice(3, 6)}`
                      ) : (
                        '------'
                      )}
                    </h2>
                    
                    <div className="flex items-center gap-3 mt-4 w-full justify-center md:justify-start">
                      <div className="w-32 h-2 bg-slate-200 rounded-full overflow-hidden relative">
                        <div 
                          className="h-full bg-gradient-to-r from-primary to-cyan-500 transition-all duration-100 ease-linear"
                          style={{ width: `${otpProgress}%` }}
                        ></div>
                      </div>
                      <span className="text-[10px] font-bold text-text-muted uppercase">Rotates in 20s</span>
                    </div>

                    <p className="text-11 text-text-secondary mt-6 max-w-xs leading-relaxed">
                      💡 Students can browse to the Sairam Portal and enter the 5-digit Access Code <strong>{activeSession.id}</strong> and the active OTP shown above.
                    </p>
                  </div>
                </div>

                {/* Action Row / Footer */}
                <div className="w-full flex flex-wrap gap-3 justify-between items-center border-t border-color pt-4">
                  <div className="flex gap-2">
                    <button 
                      onClick={() => setShowAppsScriptModal(true)}
                      className="btn-secondary py-2 px-3 text-xs rounded-lg flex items-center gap-1.5"
                    >
                      <HelpCircle size={14} /> apps script setup
                    </button>
                    <button 
                      onClick={() => setShowSettings(!showSettings)}
                      className="btn-secondary py-2 px-3 text-xs rounded-lg flex items-center gap-1.5"
                    >
                      <Settings size={14} /> custom host ip
                    </button>
                  </div>

                  <div className="flex gap-2">
                    <button 
                      onClick={handleSyncToSheets}
                      disabled={isSyncing}
                      className="btn-primary py-2 px-4 text-xs rounded-lg flex items-center gap-1.5"
                    >
                      {isSyncing ? <Loader2 size={14} className="animate-spin text-white" /> : <FileSpreadsheet size={14} className="text-white" />}
                      Sync Sheet
                    </button>
                    <button 
                      onClick={handleCloseSession}
                      className="btn-secondary border-red-500/20 hover:bg-red-500/5 text-red-400 py-2 px-4 text-xs rounded-lg"
                    >
                      Close Session
                    </button>
                  </div>
                </div>
              </div>
              
              {/* Sync Status Banner */}
              {syncStatus && (
                <div className={`p-4 rounded-xl border flex items-center justify-between slide-up ${
                  syncStatus.success 
                    ? 'badge-emerald border-emerald-500/20' 
                    : 'badge-red border-red-500/20'
                }`}>
                  <div className="flex items-center gap-2">
                    {syncStatus.success ? <CheckCircle2 size={16} className="text-emerald-400" /> : <AlertTriangle size={16} className="text-red-400" />}
                    <span className="text-xs font-bold">{syncStatus.message}</span>
                  </div>
                  <button onClick={() => setSyncStatus(null)} className="text-text-muted hover:text-text-primary">
                    <X size={14} />
                  </button>
                </div>
              )}
            </div>

            {/* Right Col: Student Lists & Controls (4 Cols) */}
            <div className="lg:col-span-4 flex flex-col gap-6">
              
              {/* Present Queue Panel */}
              <div className="glass-panel p-6 border border-white/5 flex flex-col h-350 text-left">
                <div className="flex justify-between items-center border-b border-color pb-3">
                  <div className="flex items-center gap-2">
                    <Users size={16} className="text-cyan-400" />
                    <h4 className="font-bold text-sm text-text-primary">Roster Present</h4>
                  </div>
                  <span className="badge-cyan text-xs px-2.5 py-0.5 rounded-full font-bold">
                    {activeSession.students.length} Present
                  </span>
                </div>

                {/* Present List Scrollable */}
                <div className="flex-grow overflow-y-auto mt-3 pr-1 space-y-2">
                  {activeSession.students.length === 0 ? (
                    <div className="h-full flex flex-col items-center justify-center text-center text-text-muted p-4">
                      <Loader2 size={24} className="animate-spin text-cyan-400 mb-2 opacity-50" />
                      <span className="text-xs">Waiting for students...</span>
                    </div>
                  ) : (
                    activeSession.students.map((st) => (
                      <div 
                        key={st.regNo} 
                        className="p-3 bg-slate-50 border border-color rounded-lg flex items-center justify-between text-left slide-up"
                      >
                        <div className="flex-grow">
                          <div className="text-xs font-bold text-text-primary">{st.name}</div>
                          <div className="text-10 text-text-secondary mt-0.5">{st.regNo.toUpperCase()}</div>
                          <div className="flex gap-2 mt-1.5">
                            <span className="text-9 bg-slate-200 text-text-secondary px-1.5 py-0.5 rounded font-mono">
                              {st.timestamp}
                            </span>
                            <span className={`text-9 px-1.5 py-0.5 rounded font-bold ${
                              st.type === 'OUT' 
                                ? 'bg-amber-100 text-amber-800 border border-amber-200' 
                                : 'bg-cyan-100 text-cyan-800 border border-cyan-200'
                            }`}>
                              {st.type || 'IN'}
                            </span>
                            <span className={`text-9 px-1.5 py-0.5 rounded ${
                              st.deviceStatus.includes('Shared') 
                                ? 'badge-purple font-semibold' 
                                : st.deviceStatus.includes('Manual')
                                ? 'badge-yellow font-semibold'
                                : 'badge-cyan font-semibold'
                            }`}>
                              {st.deviceStatus}
                            </span>
                          </div>
                        </div>
                        
                        <div className="flex items-center gap-2 flex-shrink-0">
                          <button 
                            type="button"
                            onClick={() => openEditStudent(st)}
                            className="text-cyan-400 hover:text-cyan-300 border-none bg-transparent cursor-pointer p-1"
                            title="Edit Student Info"
                          >
                            <Edit2 size={13} />
                          </button>
                          <CheckCircle2 size={16} className="text-emerald-400" />
                        </div>
                      </div>
                    ))
                  )}
                </div>
              </div>

              {/* Pending Approvals (Anti-Cheat Device Sharing Queue) */}
              <div className="glass-panel p-6 border border-white/5 flex flex-col h-280 text-left">
                <div className="flex items-center gap-2 border-b border-color pb-3">
                  <Smartphone size={16} className="text-purple-400" />
                  <h4 className="font-bold text-sm text-text-primary">Device Approvals</h4>
                  {activeSession.sharingRequests.length > 0 && (
                    <span className="w-2 h-2 rounded-full bg-secondary animate-ping"></span>
                  )}
                  <span className="ml-auto badge-purple text-xs px-2.5 py-0.5 rounded-full font-bold">
                    {activeSession.sharingRequests.length} Pending
                  </span>
                </div>

                <div className="flex-grow overflow-y-auto mt-3 pr-1 space-y-2">
                  {activeSession.sharingRequests.length === 0 ? (
                    <div className="h-full flex flex-col items-center justify-center text-center text-text-muted p-4">
                      <Lock size={20} className="text-purple-400/50 mb-2" />
                      <span className="text-11 leading-relaxed">No shared devices detected. Anti-cheat system monitoring is active.</span>
                    </div>
                  ) : (
                    activeSession.sharingRequests.map((req) => (
                      <div 
                        key={req.regNo}
                        className="p-3 bg-amber-500/5 border border-amber-500/10 rounded-lg slide-up text-left"
                      >
                        <div className="text-xs font-bold text-text-primary">{req.name}</div>
                        <div className="text-10 text-text-secondary mt-0.5">{req.regNo.toUpperCase()}</div>
                        <div className="text-9 badge-purple px-2 py-1 rounded mt-2 leading-relaxed">
                          ⚠️ Device already logged attendance for: <strong className="font-bold">{req.originalRegNo.toUpperCase()}</strong>
                        </div>
                        
                        <div className="flex gap-2 mt-3 justify-end">
                          <button 
                            onClick={() => handleRejectSharing(req.regNo)}
                            className="bg-slate-100 hover:bg-slate-200 text-text-secondary text-10 font-bold px-3 py-1.5 rounded-md border border-color"
                          >
                            Reject
                          </button>
                          <button 
                            onClick={() => handleApproveSharing(req.regNo)}
                            className="btn-primary py-1.5 px-3 text-10 font-bold rounded-md shadow-md"
                          >
                            Approve Exception
                          </button>
                        </div>
                      </div>
                    ))
                  )}
                </div>
              </div>

              {/* Manual Check-in Board */}
              <div className="glass-panel p-6 border border-white/5 text-left">
                <h4 className="font-bold text-sm border-b border-color pb-3 flex items-center gap-2 text-text-primary">
                  <Plus size={16} className="text-yellow-400" />
                  Manual ERP Entry
                </h4>
                <form onSubmit={handleManualMark} className="mt-3 space-y-3 text-left">
                  <div>
                    <label className="block text-10 font-semibold text-text-secondary uppercase mb-1">Student ID</label>
                    <div className="flex gap-2">
                      <input 
                        type="text" 
                        placeholder="e.g. STUDENT123" 
                        value={searchRegNo} 
                        onChange={e => setSearchRegNo(e.target.value.toLowerCase())}
                        className="glass-input py-1.5 text-xs" 
                        required 
                      />
                      <button 
                        type="submit" 
                        className="btn-primary py-1.5 px-4 text-xs rounded-lg"
                      >
                        Check-in
                      </button>
                    </div>
                  </div>
                </form>
              </div>

            </div>

          </div>
        )}

        {/* ==========================================
            VIEW 4: STUDENT MARKING PORTAL
            ========================================== */}
        {view === 'student' && (
          <div className="w-full max-w-md py-4 slide-up text-left">
            
            {/* Geofence Radar Card */}
            <div className="glass-panel p-6 border border-white/5 flex items-center gap-4 mb-4">
              <div className="flex-shrink-0">
                <div className={`geofence-radar ${
                  geoStatus === 'success' 
                    ? 'border-emerald-500/20' 
                    : geoStatus === 'out_of_bounds' || geoStatus === 'denied' 
                    ? 'border-red-500/20' 
                    : 'border-cyan-500/20'
                }`}>
                  <div className={`radar-dot ${
                    geoStatus === 'success' 
                      ? 'bg-emerald-500 shadow-emerald-500' 
                      : geoStatus === 'out_of_bounds' || geoStatus === 'denied' 
                      ? 'bg-red-500 shadow-red-500' 
                      : 'bg-cyan-600 shadow-lg'
                  }`}></div>
                </div>
              </div>

              <div>
                <h4 className="font-bold text-sm text-text-primary">Classroom GPS Verification</h4>
                
                {geoStatus === 'idle' && (
                  <p className="text-xs text-text-muted mt-1">Awaiting coordinates verification...</p>
                )}
                {geoStatus === 'checking' && (
                  <div className="flex items-center gap-1 text-xs text-cyan-400 mt-1 font-bold">
                    <Loader2 size={12} className="animate-spin text-cyan-400" /> Computing distance to campus...
                  </div>
                )}
                {geoStatus === 'success' && (
                  <div>
                    <p className="text-xs text-emerald-400 font-bold mt-1">✓ Coordinates Verified Inside Campus</p>
                    <p className="text-10 text-text-muted mt-0.5">
                      Distance: {studentDistance !== null ? `${Math.round(studentDistance)} meters` : 'Verified'}
                    </p>
                  </div>
                )}
                {geoStatus === 'out_of_bounds' && (
                  <div>
                    <p className="text-xs text-red-400 font-bold mt-1">⚠️ Out of Geofence Boundary</p>
                    <p className="text-10 text-text-muted mt-0.5 leading-normal">
                      Distance: {studentDistance !== null ? `${Math.round(studentDistance)}m` : 'Unknown'}. You must stay inside Sri Sairam campus.
                    </p>
                  </div>
                )}
                {geoStatus === 'denied' && (
                  <div>
                    <p className="text-xs text-red-400 font-bold mt-1">❌ Location Access Denied</p>
                    <p className="text-10 text-text-muted mt-0.5 leading-normal">
                      Browser GPS permissions are disabled. Please enable location services in your browser settings to verify.
                    </p>
                  </div>
                )}

                <button 
                  onClick={requestStudentLocation}
                  className="text-10 font-bold text-cyan-400 flex items-center gap-1 mt-2.5 hover:underline"
                >
                  <RefreshCw size={10} /> Recalculate Distance
                </button>
              </div>
            </div>

            {/* Main Form Card */}
            <div className="glass-panel p-8 border border-white/5">
              
              <div className="border-b border-color pb-4 mb-6">
                <h3 className="text-xl font-bold text-text-primary">Student Sign-in</h3>
                <p className="text-xs text-text-secondary mt-1">Provide your credentials and current session verification keys</p>
              </div>

              {attendanceError && (
                <div className="p-3 bg-red-500/5 border border-red-500/15 text-red-400 text-xs rounded-lg mb-4 flex items-start gap-2">
                  <AlertTriangle size={14} className="flex-shrink-0 mt-0.5 text-red-400" />
                  <span>{attendanceError}</span>
                </div>
              )}

              {/* Pending approval state loader */}
              {pendingApproval ? (
                <div className="text-center py-8 space-y-4">
                  <Loader2 size={36} className="animate-spin text-purple-400 mx-auto" />
                  <h4 className="text-lg font-bold text-purple-400">Exception Approval Pending</h4>
                  <p className="text-xs text-text-secondary max-w-xs mx-auto leading-relaxed">
                    This phone was already used to mark attendance for another student. We have sent an exception request to your teacher.
                  </p>
                  <div className="p-4 bg-slate-50 border border-color rounded-xl text-left text-xs max-w-xs mx-auto space-y-2">
                    <div><span className="text-text-secondary">Student ID:</span> <span className="font-bold text-text-primary">{studentRegNo.toUpperCase()}</span></div>
                    <div><span className="text-text-secondary">Device ID:</span> <span className="font-mono text-text-muted">{studentDeviceFingerprint}</span></div>
                    <div><span className="text-text-secondary">Status:</span> <span className="text-yellow-400 font-bold">Waiting for teacher approval...</span></div>
                  </div>
                  <p className="text-10 text-text-muted leading-relaxed">
                    Your screen will automatically redirect to success once approved on the Smart TV screen.
                  </p>
                </div>
              ) : (
                <form onSubmit={handleStudentSubmit} className="space-y-4">
                  
                   {/* Session Id (if not auto loaded) */}
                  {!scannedToken && (
                    <div>
                      <button 
                        type="button" 
                        onClick={startScanner}
                        className="btn-secondary w-full py-2.5 mb-4 text-xs flex items-center justify-center gap-2"
                      >
                        <Camera size={14} /> Scan Smart TV QR Code
                      </button>

                      <label className="block text-xs font-semibold text-text-secondary uppercase mb-1">Class Access Code</label>
                      <input 
                        type="text" 
                        maxLength={10}
                        placeholder="e.g. G3103" 
                        value={classCode} 
                        onChange={e => setClassCode(e.target.value.toUpperCase().replace(/[^A-Z0-9-]/g, ''))}
                        className="glass-input text-center font-bold tracking-widest text-lg font-mono" 
                        required 
                      />
                      
                      {isClassSearching && (
                        <div className="flex items-center gap-1.5 text-xs text-cyan-400 font-semibold mt-1.5">
                          <Loader2 size={12} className="animate-spin text-cyan-400" /> Searching class details...
                        </div>
                      )}
                      
                      {classLookupError && (
                        <p className="text-xs font-bold text-red-400 mt-1.5">{classLookupError}</p>
                      )}

                      {classLookup && (
                        <div className="p-3 bg-emerald-500/5 border border-emerald-500/10 text-emerald-400 text-xs rounded-xl mt-2 flex flex-col gap-1 text-left slide-up">
                          <div className="font-bold text-text-primary">✓ Class Found</div>
                          <div><span className="text-text-secondary">Classroom Code:</span> <span className="font-semibold">{classLookup.classroomCode}</span></div>
                          <div><span className="text-text-secondary">Instructor:</span> <span className="font-semibold">{classLookup.teacherName}</span></div>
                          <div><span className="text-text-secondary">Target Roster:</span> <span className="font-semibold">{classLookup.year} Year {classLookup.department} - Sec {classLookup.section}</span></div>
                        </div>
                      )}
                    </div>
                  )}

                  {/* Student ID Input */}
                  <div>
                    <label className="block text-xs font-semibold text-text-secondary uppercase mb-1">Student ID</label>
                    <input 
                      type="text" 
                      placeholder="e.g. STUDENT123" 
                      value={studentRegNo} 
                      onChange={e => setStudentRegNo(e.target.value.toLowerCase())}
                      className="glass-input uppercase" 
                      required 
                    />
                    {studentMatchingMsg && (
                      <p className={`text-xs font-bold mt-1.5 ${studentMatchingMsg.includes('❌') ? 'text-red-400' : 'text-cyan-400'}`}>{studentMatchingMsg}</p>
                    )}
                  </div>

                  {/* OTP Validation code */}
                  <div>
                    <label className="block text-xs font-semibold text-text-secondary uppercase mb-1">TV Verification OTP</label>
                    {scannedToken ? (
                      <div className="flex items-center gap-2 p-3 badge-cyan rounded-xl">
                        <Lock size={14} className="text-cyan-400" />
                        <span className="text-xs font-bold font-mono">Dynamic Token Loaded via QR</span>
                      </div>
                    ) : (
                      <input 
                        type="text" 
                        maxLength={6}
                        placeholder="Enter the 6-digit OTP code" 
                        value={studentOtp} 
                        onChange={e => setStudentOtp(e.target.value.replace(/\D/g, ''))}
                        className="glass-input text-center tracking-widest text-lg font-bold font-mono" 
                        required 
                      />
                    )}
                  </div>

                  {/* Device Info */}
                  <div className="pt-2">
                    <span className="text-10 text-text-muted block text-right font-mono">
                      Device Token: {studentDeviceFingerprint} ({studentDeviceDesc})
                    </span>
                  </div>

                  <button 
                    type="submit" 
                    disabled={isSubmitting || geoStatus !== 'success'}
                    className="btn-primary w-full py-3.5 mt-4"
                  >
                    {isSubmitting ? (
                      <span className="flex items-center gap-2">
                        <Loader2 size={16} className="animate-spin text-white" /> Logging present status...
                      </span>
                    ) : (
                      <span className="flex items-center gap-2 text-white">
                        <Send size={16} className="text-white" /> Submit Attendance Present
                      </span>
                    )}
                  </button>

                </form>
              )}

            </div>
          </div>
        )}

        {/* ==========================================
            VIEW 5: STUDENT ATTENDANCE SUCCESS
            ========================================== */}
        {view === 'student_success' && markedStudentDetails && (
          <div className="w-full max-w-md py-8 slide-up text-center">
            <div className="glass-panel p-8 border border-white/5 relative overflow-hidden">
              
              <div className="w-16 h-16 rounded-full bg-emerald-500/10 border border-emerald-500/20 flex items-center justify-center text-emerald-400 mx-auto mb-6 scale-110 pulsing-ring">
                <CheckCircle2 size={32} className="text-emerald-400" />
              </div>

              <span className="text-10 font-bold badge-emerald px-3 py-1 rounded-full uppercase">
                Verified Present
              </span>

              <h2 className="text-3xl font-black text-text-primary mt-6 tracking-tight">Attendance Logged</h2>
              <p className="text-xs text-text-secondary mt-1">Your presence has been successfully verified and registered in the ERP</p>

              <div className="my-8 p-4 bg-slate-50 border border-color rounded-xl text-left text-xs space-y-2.5">
                <div className="flex justify-between border-b border-color pb-2">
                  <span className="text-text-secondary">Student Name</span>
                  <span className="font-bold text-text-primary">{markedStudentDetails.name}</span>
                </div>
                <div className="flex justify-between border-b border-color pb-2">
                  <span className="text-text-secondary">Student ID</span>
                  <span className="font-bold text-text-primary font-mono uppercase">{markedStudentDetails.regNo}</span>
                </div>
                <div className="flex justify-between border-b border-color pb-2">
                  <span className="text-text-secondary">Sign-in Time</span>
                  <span className="font-bold text-text-primary">{markedStudentDetails.timestamp}</span>
                </div>
                <div className="flex justify-between border-b border-color pb-2">
                  <span className="text-text-secondary">Verification Method</span>
                  <span className="font-bold text-cyan-400">{markedStudentDetails.method}</span>
                </div>
                <div className="flex justify-between">
                  <span className="text-text-secondary">Device Trust Level</span>
                  <span className="font-bold text-purple-400">{markedStudentDetails.deviceStatus}</span>
                </div>
              </div>

              <div className="flex gap-3">
                <button 
                  onClick={() => {
                    setView('landing');
                    setClassCode('');
                    setStudentRegNo('');
                    setStudentOtp('');
                    setScannedToken('');
                    setGeoStatus('idle');
                  }}
                  className="btn-secondary w-full py-3"
                >
                  Done & Close Portal
                </button>
              </div>
            </div>
          </div>
        )}

      </main>

      {/* Footer */}
      <footer className="py-4 text-center text-text-muted text-10 font-semibold tracking-wider uppercase border-t border-color max-w-7xl mx-auto w-full">
        © 2026 Sri Sairam Engineering College. Optimized for Classroom Smart TV & Student Mobile Browser.
      </footer>

      {/* ==========================================
          APPS SCRIPT SETUP GUIDE MODAL
          ========================================== */}
      {showAppsScriptModal && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-filter backdrop-blur-md flex items-center justify-center p-4 z-50 animate-fade-in text-left">
          <div className="glass-panel p-6 max-w-2xl w-full border border-color shadow-2xl bg-white flex flex-col max-h-85vh relative">
            
            {/* Fixed Header */}
            <div className="flex justify-between items-start border-b border-color pb-4 mb-4 relative pr-10">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 rounded-lg bg-emerald-500/10 border border-emerald-500/20 flex items-center justify-center text-emerald-400 flex-shrink-0">
                  <FileSpreadsheet size={20} className="text-emerald-400" />
                </div>
                <div>
                  <h3 className="text-lg font-bold text-text-primary">Google Sheets Setup Instructions</h3>
                  <p className="text-xs text-text-secondary">Follow these steps to link your attendance board to a sheet</p>
                </div>
              </div>
              
              <button 
                type="button"
                onClick={() => setShowAppsScriptModal(false)}
                className="absolute top-0 right-0 text-text-secondary hover:text-text-primary border-none bg-transparent cursor-pointer hover:bg-slate-100 p-1.5 rounded-full transition-colors z-30 flex items-center justify-center"
              >
                <X size={20} />
              </button>
            </div>

            {/* Scrollable Content Area */}
            <div className="flex-grow overflow-y-auto space-y-4 text-xs text-text-secondary leading-relaxed pr-1 h-280 scrollbar">
              <p>
                To automatically feed attendance logs into your personal Google Sheet, complete these simple 2-minute setup steps:
              </p>
              
              <ol className="list-decimal list-inside space-y-2 text-[11px]">
                <li>Create a new Google Sheet (or open an existing one) where you want to write the student list.</li>
                <li>Go to the top menu and select <strong>Extensions &gt; Apps Script</strong>.</li>
                <li>Delete any default code in the editor and paste the code template provided below.</li>
                <li>Click the 💾 (Save) icon, then click <strong>Deploy &gt; New deployment</strong>.</li>
                <li>Select type: <strong>Web app</strong>. Change "Who has access" to <strong>Anyone</strong> (this is required so the local server can post coordinates).</li>
                <li>Click <strong>Deploy</strong>, authorize the Google permissions, and copy the generated <strong>Web App URL</strong>.</li>
                <li>Paste that URL into the <strong>Google Apps Script URL</strong> field in this application settings.</li>
              </ol>

              <div className="mt-4">
                <div className="flex justify-between items-center bg-slate-50 border border-color border-b-0 rounded-t-lg px-4 py-2">
                  <span className="font-bold text-10 text-text-primary">Google Apps Script Code</span>
                  <button 
                    onClick={() => {
                      navigator.clipboard.writeText(APPS_SCRIPT_DOPOST);
                      alert("Google Apps Script code copied to clipboard!");
                    }}
                    className="flex items-center gap-1 hover:text-text-primary border-none bg-transparent cursor-pointer font-bold text-10 text-cyan-400"
                  >
                    <Clipboard size={12} /> Copy Code
                  </button>
                </div>
                <pre className="bg-[#0f172a] border border-color p-3 rounded-b-lg font-mono text-[9px] overflow-auto text-slate-300 h-48 scrollbar text-left">
                  {APPS_SCRIPT_DOPOST}
                </pre>
              </div>
            </div>

            {/* Fixed Footer */}
            <div className="border-t border-color pt-4 mt-4 flex justify-end">
              <button 
                type="button"
                onClick={() => setShowAppsScriptModal(false)}
                className="btn-primary w-full py-2.5 rounded-lg text-xs"
              >
                Close Setup Instructions
              </button>
            </div>

          </div>
        </div>
      )}

      {/* ==========================================
          CAMERA QR SCANNER VIEWFINDER MODAL
          ========================================== */}
      {showScanner && (
        <div className="scanner-modal">
          <div className="scanner-viewfinder slide-up">
            <div className="flex justify-between items-center border-b border-color pb-3 mb-3">
              <span className="font-bold text-sm text-text-primary">Camera QR Scan</span>
              <button 
                type="button" 
                onClick={stopScanner}
                className="text-text-muted hover:text-text-primary border-none bg-transparent cursor-pointer"
              >
                <X size={18} />
              </button>
            </div>
            
            {scannerError && (
              <div className="p-3 bg-red-500/5 border border-red-500/15 text-red-400 text-xs rounded-lg mb-3 flex items-start gap-2 text-left">
                <AlertTriangle size={14} className="flex-shrink-0 mt-0.5 text-red-400" />
                <span>{scannerError}</span>
              </div>
            )}
            
            <div className="scanner-box-container">
              <div className="scanner-laser"></div>
              <div id="qr-reader"></div>
            </div>

            {/* Zoom Slider and Buttons controls */}
            {zoomSupported && (
              <div className="mt-4 px-2 space-y-2 text-left border border-white/5 bg-slate-500/5 p-3 rounded-xl">
                <div className="flex justify-between items-center text-[10px] font-bold text-text-secondary uppercase">
                  <span>Camera Zoom ({currentZoom.toFixed(1)}x)</span>
                  <div className="flex gap-1">
                    {[1, 2, 3].map(z => (
                      <button
                        key={z}
                        type="button"
                        onClick={() => handleZoomChange(Math.min(maxZoom, Math.max(minZoom, z)))}
                        className={`px-2 py-0.5 text-[9px] rounded font-bold transition-all border ${
                          Math.abs(currentZoom - z) < 0.1
                            ? 'bg-cyan-400 border-cyan-400 text-slate-900'
                            : 'bg-slate-800/40 border-color text-text-secondary hover:text-text-primary'
                        }`}
                      >
                        {z}x
                      </button>
                    ))}
                    <button
                      type="button"
                      onClick={() => handleZoomChange(maxZoom)}
                      className={`px-2 py-0.5 text-[9px] rounded font-bold transition-all border ${
                        Math.abs(currentZoom - maxZoom) < 0.1
                          ? 'bg-cyan-400 border-cyan-400 text-slate-900'
                          : 'bg-slate-800/40 border-color text-text-secondary hover:text-text-primary'
                      }`}
                    >
                      Max
                    </button>
                  </div>
                </div>
                <input 
                  type="range"
                  min={minZoom}
                  max={maxZoom}
                  step={0.1}
                  value={currentZoom}
                  onChange={e => handleZoomChange(parseFloat(e.target.value))}
                  className="w-full h-1.5 bg-slate-800/40 rounded-lg appearance-none cursor-pointer accent-cyan-400 border border-color"
                />
              </div>
            )}
            
            <p className="text-11 text-text-secondary mt-3">
              Point your camera at the QR code displayed on the classroom screen.
            </p>
            
            <button 
              type="button" 
              onClick={stopScanner}
              className="btn-secondary w-full py-2 rounded-lg mt-4 text-xs"
            >
              Cancel Scanning
            </button>
          </div>
        </div>
      )}

      {/* ==========================================
          NETWORK SETTINGS MODAL (Exposed Host Configuration)
          ========================================== */}
      {showSettings && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-filter backdrop-blur-md flex items-center justify-center p-4 z-50 animate-fade-in text-left">
          <div className="glass-panel p-6 border border-color shadow-2xl max-w-sm w-full bg-white relative">
            
            <button 
              onClick={() => setShowSettings(false)}
              className="absolute top-4 right-4 text-text-secondary hover:text-text-primary border-none bg-transparent cursor-pointer"
            >
              <X size={20} />
            </button>

            <div className="flex items-center gap-2 border-b border-color pb-3 mb-4">
              <Settings size={18} className="text-cyan-400" />
              <h4 className="text-sm font-bold text-text-primary">Network Settings</h4>
            </div>
            
            <div className="space-y-4">
              <div>
                <label className="text-10 text-text-secondary uppercase font-semibold block mb-1">Wi-Fi Host IP Address</label>
                <input 
                  type="text" 
                  placeholder="e.g. 192.168.0.107" 
                  value={teacherCustomIp} 
                  onChange={e => setTeacherCustomIp(e.target.value)}
                  className="glass-input text-sm"
                />
              </div>
              
              <p className="text-9 text-text-muted leading-normal">
                Students scanning from phones must connect to the same local network. Enter your laptop's current local Wi-Fi IP address. 
                <br />
                <span className="font-semibold text-text-secondary">Current server host: {window.location.hostname}</span>
              </p>
            </div>

            <button 
              onClick={() => setShowSettings(false)}
              className="btn-primary w-full py-2.5 rounded-lg mt-5 text-xs font-bold shadow-md"
            >
              Save IP Configuration
            </button>
          </div>
        </div>
      )}

      {/* ==========================================
          ROSTER DATABASE SHEET GUIDE MODAL
          ========================================== */}
      {showRosterGuideModal && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-filter backdrop-blur-md flex items-center justify-center p-4 z-50 animate-fade-in text-left">
          <div className="glass-panel p-6 border border-color shadow-2xl max-w-xl w-full bg-white relative flex flex-col max-h-85vh">
            
            {/* Fixed Header */}
            <div className="flex justify-between items-center border-b border-color pb-3 mb-4">
              <div className="flex items-center gap-2">
                <FileSpreadsheet size={20} className="text-purple-400" />
                <h3 className="text-base font-black text-text-primary">Google Sheet Roster Database Guide</h3>
              </div>
              <button 
                type="button" 
                onClick={() => setShowRosterGuideModal(false)}
                className="text-text-muted hover:text-text-primary border-none bg-transparent cursor-pointer"
              >
                <X size={20} />
              </button>
            </div>

            {/* Scrollable Body */}
            <div className="overflow-auto pr-2 space-y-4 text-xs text-text-secondary leading-relaxed h-280 flex-grow">
              <p>
                To eliminate mock accounts, format a new Google Sheet to host your college's live student roster and faculty credentials. Follow these steps:
              </p>

              <h4 className="font-bold text-text-primary uppercase tracking-wider text-10 mt-3 border-b border-color pb-1">1. Format Spreadsheet Tabs</h4>
              <ul className="list-disc pl-4 space-y-1 text-10">
                <li>Create a sheet tab named <strong className="text-text-primary">Faculty</strong>. Add headers: <code className="font-mono text-purple-400 font-bold">Username</code> | <code className="font-mono text-purple-400 font-bold">Password</code> | <code className="font-mono text-purple-400 font-bold">Name</code> | <code className="font-mono text-purple-400 font-bold">Department</code></li>
                <li>Create a tab named <strong className="text-text-primary">Students</strong>. Add headers: <code className="font-mono text-purple-400 font-bold">Student ID</code> | <code className="font-mono text-purple-400 font-bold">Student Name</code> | <code className="font-mono text-purple-400 font-bold">Department</code> | <code className="font-mono text-purple-400 font-bold">Section</code> | <code className="font-mono text-purple-400 font-bold">Year</code></li>
                <li>Create a tab named <strong className="text-text-primary">Classrooms</strong>. Add headers: <code className="font-mono text-purple-400 font-bold">Class Code</code> | <code className="font-mono text-purple-400 font-bold">Latitude</code> | <code className="font-mono text-purple-400 font-bold">Longitude</code> | <code className="font-mono text-purple-400 font-bold">Radius</code></li>
              </ul>

              <h4 className="font-bold text-text-primary uppercase tracking-wider text-10 mt-3 border-b border-color pb-1">2. Add Apps Script Code</h4>
              <p>Click <strong className="text-text-primary">Extensions &gt; Apps Script</strong> in the sheet menu, replace all existing code with the script below, and save:</p>

              <div className="mt-2 relative">
                <div className="flex justify-between items-center bg-slate-50 border border-color border-b-0 rounded-t-lg px-4 py-2">
                  <span className="font-bold text-10 text-text-primary">Roster DB Script Code</span>
                  <button 
                    onClick={() => {
                      navigator.clipboard.writeText(APPS_SCRIPT_DOGET);
                      alert("Roster DB script copied to clipboard!");
                    }}
                    className="text-10 font-bold text-cyan-400 hover:text-cyan-300 border-none bg-transparent cursor-pointer flex items-center gap-1"
                  >
                    <Clipboard size={12} /> Copy Code
                  </button>
                </div>
                <pre className="bg-[#0f172a] border border-color p-3 rounded-b-lg font-mono text-[9px] overflow-auto text-slate-300 h-40 scrollbar text-left">
                  {APPS_SCRIPT_DOGET}
                </pre>
              </div>

              <h4 className="font-bold text-text-primary uppercase tracking-wider text-10 mt-3 border-b border-color pb-1">3. Deploy Web App</h4>
              <ol className="list-decimal pl-4 space-y-1 text-10">
                <li>In Apps Script editor, click <strong className="text-text-primary">Deploy &gt; New Deployment</strong>.</li>
                <li>Select type: <strong className="text-text-primary">Web App</strong>.</li>
                <li>Execute as: <strong className="text-text-primary">Me</strong>.</li>
                <li>Who has access: <strong className="text-text-primary">Anyone</strong>.</li>
                <li>Click <strong className="text-text-primary">Deploy</strong>, authorize permissions, and copy the Web App URL (ends in <code className="font-mono text-purple-400">/exec</code>).</li>
              </ol>
            </div>

            {/* Fixed Footer */}
            <div className="border-t border-color pt-4 mt-4 flex justify-end">
              <button 
                type="button"
                onClick={() => setShowRosterGuideModal(false)}
                className="btn-primary w-full py-2.5 rounded-lg text-xs"
              >
                Close Setup Instructions
              </button>
            </div>

          </div>
        </div>
      )}

      {/* ==========================================
          CHANGE PASSWORD MODAL
          ========================================== */}
      {showChangePasswordModal && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-filter backdrop-blur-md flex items-center justify-center p-4 z-50 animate-fade-in text-left">
          <div className="glass-panel p-6 border border-color shadow-2xl max-w-sm w-full bg-white relative">
            
            <button 
              onClick={() => setShowChangePasswordModal(false)}
              className="absolute top-4 right-4 text-text-secondary hover:text-text-primary border-none bg-transparent cursor-pointer"
            >
              <X size={20} />
            </button>

            <div className="flex items-center gap-2 border-b border-color pb-3 mb-4">
              <Lock size={18} className="text-cyan-400" />
              <h4 className="text-sm font-bold text-text-primary">Change Faculty Password</h4>
            </div>
            
            <form onSubmit={handlePasswordChange} className="space-y-4">
              <div>
                <label className="text-10 text-text-secondary uppercase font-semibold block mb-1">Current Password</label>
                <input 
                  type="password" 
                  required
                  placeholder="Enter current password" 
                  value={currentPassword} 
                  onChange={e => setCurrentPassword(e.target.value)}
                  className="glass-input text-xs"
                />
              </div>

              <div>
                <label className="text-10 text-text-secondary uppercase font-semibold block mb-1">New Password</label>
                <input 
                  type="password" 
                  required
                  placeholder="Enter new password" 
                  value={newPassword} 
                  onChange={e => setNewPassword(e.target.value)}
                  className="glass-input text-xs"
                />
              </div>

              <div>
                <label className="text-10 text-text-secondary uppercase font-semibold block mb-1">Confirm New Password</label>
                <input 
                  type="password" 
                  required
                  placeholder="Confirm new password" 
                  value={confirmNewPassword} 
                  onChange={e => setConfirmNewPassword(e.target.value)}
                  className="glass-input text-xs"
                />
              </div>

              {changePasswordMsg && (
                <div className={`p-2 rounded text-[10px] font-bold ${
                  changePasswordMsg.startsWith('✓') 
                    ? 'bg-emerald-500/5 text-emerald-400 border border-emerald-500/10'
                    : 'bg-red-500/5 text-red-400 border border-red-500/10'
                }`}>
                  {changePasswordMsg}
                </div>
              )}

              <button 
                type="submit"
                disabled={isChangingPassword}
                className="btn-primary w-full py-2.5 rounded-lg mt-2 text-xs font-bold shadow-md animate-pulse"
              >
                {isChangingPassword ? "Updating Password..." : "Update Password"}
              </button>
            </form>
          </div>
        </div>
      )}

      {/* ==========================================
          EDIT STUDENT DETAILS MODAL
          ========================================== */}
      {showEditStudentModal && editingStudent && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-filter backdrop-blur-md flex items-center justify-center p-4 z-50 animate-fade-in text-left">
          <div className="glass-panel p-6 border border-color shadow-2xl max-w-sm w-full bg-white relative">
            
            <button 
              onClick={() => {
                setShowEditStudentModal(false);
                setEditingStudent(null);
              }}
              className="absolute top-4 right-4 text-text-secondary hover:text-text-primary border-none bg-transparent cursor-pointer"
            >
              <X size={20} />
            </button>

            <div className="flex items-center gap-2 border-b border-color pb-3 mb-4">
              <Edit2 size={18} className="text-cyan-400" />
              <h4 className="text-sm font-bold text-text-primary">Edit Student Details</h4>
            </div>
            
            <form onSubmit={handleSaveEditStudent} className="space-y-4">
              <div>
                <label className="block text-10 font-semibold text-text-secondary uppercase mb-1">Student Name</label>
                <input 
                  type="text" 
                  required
                  placeholder="e.g. Student Name" 
                  value={editStudentName} 
                  onChange={e => setEditStudentName(e.target.value)}
                  className="glass-input text-xs"
                />
              </div>



              <div>
                <label className="block text-10 font-semibold text-text-secondary uppercase mb-1">Student ID</label>
                <input 
                  type="text" 
                  required
                  placeholder="e.g. SEC23EC242" 
                  value={editStudentReg} 
                  onChange={e => setEditStudentReg(e.target.value.toLowerCase())}
                  className="glass-input uppercase text-xs"
                />
              </div>

              <div className="flex gap-2 pt-2">
                <button 
                  type="button"
                  onClick={() => {
                    setShowEditStudentModal(false);
                    setEditingStudent(null);
                  }}
                  className="btn-secondary w-full py-2.5 rounded-lg text-xs"
                >
                  Cancel
                </button>
                <button 
                  type="submit"
                  className="btn-primary w-full py-2.5 rounded-lg text-xs font-bold shadow-md"
                >
                  Save Changes
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

    </div>
  );
}
