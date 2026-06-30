/**
 * Generates a stable and unique fingerprint for the student's browser device.
 * Used to identify when multiple students are attempting to log in or mark attendance from the same device.
 */
export function getDeviceFingerprint(): string {
  // 1. Persistent UUID in localStorage to survive restarts
  let deviceUuid = localStorage.getItem("sairam_attendance_device_uuid");
  if (!deviceUuid) {
    // Generate a unique device UUID
    deviceUuid = 'dev_' + 
      Math.random().toString(36).substring(2, 15) + 
      Math.random().toString(36).substring(2, 15);
    localStorage.setItem("sairam_attendance_device_uuid", deviceUuid);
  }

  // 2. Hardware and Browser characteristics to bind the token
  const screenSpec = `${window.screen.width}x${window.screen.height}x${window.screen.colorDepth}`;
  const userAgent = navigator.userAgent;
  const language = navigator.language || "en";
  const timeZone = Intl.DateTimeFormat().resolvedOptions().timeZone || "UTC";

  const rawFingerprint = `${deviceUuid}|${screenSpec}|${userAgent}|${language}|${timeZone}`;

  // 3. Generate a fast numeric hash to act as a readable fingerprint ID
  let hash = 0;
  for (let i = 0; i < rawFingerprint.length; i++) {
    const char = rawFingerprint.charCodeAt(i);
    hash = (hash << 5) - hash + char;
    hash |= 0; // Convert to a 32-bit signed integer
  }

  // Returns a readable fingerprint code (e.g. FPR-1928374-f8b1c2)
  const shortUuid = deviceUuid.replace("dev_", "").substring(0, 6);
  return `FPR-${Math.abs(hash).toString(16).toUpperCase()}-${shortUuid}`;
}

/**
 * Returns basic readable device information.
 */
export function getDeviceDescription(): string {
  const userAgent = navigator.userAgent;
  let deviceType = "Mobile Device";
  
  if (userAgent.match(/iPad|iPhone|iPod/i)) {
    deviceType = "iOS Device";
  } else if (userAgent.match(/Android/i)) {
    deviceType = "Android Device";
  } else if (userAgent.match(/Macintosh/i)) {
    deviceType = "macOS Desktop";
  } else if (userAgent.match(/Windows/i)) {
    deviceType = "Windows PC";
  } else if (userAgent.match(/Linux/i)) {
    deviceType = "Linux System";
  }

  return deviceType;
}
