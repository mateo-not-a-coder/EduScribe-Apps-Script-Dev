/**
 * ======================================================================
 * EduScribe-DEV - V9 - Gemini Transcription Workflow
 * ======================================================================
 * Apps Script handles Drive monitoring and triggers the Cloud Run service.
 * Cloud Run handles the entire transcription workflow.
 * This script also contains utilities for post-processing and homework.
 */

// ------------------------------
// GLOBAL CONFIG (LOADED FROM PropertiesService)
// ------------------------------
// FOLDER_ID                  (Drive folder for new Meet recordings)
// TRACKING_SHEET_ID          (ID of Spreadsheet for job tracking)
// CLOUD_RUN_URL              (URL of the Cloud Run service)
// STUDENT_ROSTER_ID          (ID of Spreadsheet containing 'Current_Students' etc.)
// HOMEWORK_SHEET_NAME        (Name of the homework ledger tab, e.g., "Homework_Push")
// HOMEWORK_PORTAL_BASEURL    (Base URL for the homework portal link)
// ROSTER_DRIVE_FOLDER_ID     (0-based index of the Drive_Folder_ID column)
// CLIENT_EMAIL & PRIVATE_KEY (For getServiceAccountToken, if needed by helpers)

/**
 * Retrieves a short-lived OAuth2 access token for the service account.
 * NOTE: This is no longer used by the primary transcription workflow, but may
 * be useful for other utility or testing functions.
 */
function getServiceAccountToken() {
  try {
    const localScriptProperties = PropertiesService.getScriptProperties();
    const privateKeyProperty = localScriptProperties.getProperty('PRIVATE_KEY');
    const clientEmail = localScriptProperties.getProperty('CLIENT_EMAIL');
    if (!privateKeyProperty || !clientEmail) {
      throw new Error("CRITICAL: Missing PRIVATE_KEY or CLIENT_EMAIL in Script Properties.");
    }
    const privateKey = privateKeyProperty
      .replace(/\\n/g, '\n')
      .replace("-----BEGIN PRIVATE KEY----- ", "-----BEGIN PRIVATE KEY-----\n")
      .replace(" -----END PRIVATE KEY-----", "\n-----END PRIVATE KEY-----");

    const now = Math.floor(Date.now() / 1000);
    const expiration = now + 3600;
    const header = JSON.stringify({ alg: 'RS256', typ: 'JWT' });
    const encodedHeader = Utilities.base64EncodeWebSafe(header);
    const claimSet = JSON.stringify({
      iss: clientEmail,
      scope: 'https://www.googleapis.com/auth/devstorage.full_control',
      aud: 'https://oauth2.googleapis.com/token',
      iat: now,
      exp: expiration
    });
    const encodedClaimSet = Utilities.base64EncodeWebSafe(claimSet);
    const signatureInput = `${encodedHeader}.${encodedClaimSet}`;
    const signatureBytes = Utilities.computeRsaSha256Signature(signatureInput, privateKey);
    const encodedSignature = Utilities.base64EncodeWebSafe(signatureBytes);
    const assertion = `${signatureInput}.${encodedSignature}`;

    const tokenResponse = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
      method: 'post',
      contentType: 'application/x-www-form-urlencoded',
      payload: { grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer', assertion: assertion },
      muteHttpExceptions: true
    });

    const responseCode = tokenResponse.getResponseCode();
    const responseText = tokenResponse.getContentText();
    if (responseCode !== 200) {
      throw new Error(`Failed to fetch token from Google (${responseCode}).`);
    }
    const tokenData = JSON.parse(responseText);
    if (!tokenData.access_token) {
      throw new Error("Access token not found in Google's response.");
    }
    return tokenData.access_token;
  } catch (e) {
    Logger.log(`[AUTH ERROR] Unexpected exception in getServiceAccountToken: ${e.message}\nStack: ${e.stack}`);
    throw new Error(`[AUTH ERROR] Authentication failed: ${e.message}`);
  }
}

// ------------------------------
// PRIMARY TRIGGER AND HELPERS
// ------------------------------

/**
 * Finds new video recordings in Drive, renames them, calls Cloud Run
 * to accept the file for processing, and updates the tracking sheet.
 * This is the main trigger function for the entire workflow.
 */
function processNewRecordings() {
  const props = PropertiesService.getScriptProperties();
  const currentFolderId = props.getProperty("FOLDER_ID");
  const currentTrackingSheetId = props.getProperty("TRACKING_SHEET_ID");
  const currentCloudRunUrl = props.getProperty("CLOUD_RUN_URL");

  if (!currentFolderId) { Logger.log("ERROR: FOLDER_ID script property is not set."); return; }
  if (!currentTrackingSheetId) { Logger.log("ERROR: TRACKING_SHEET_ID script property is not set."); return; }
  if (!currentCloudRunUrl) { Logger.log("ERROR: CLOUD_RUN_URL script property is not set."); return; }

  try {
    const folder = DriveApp.getFolderById(currentFolderId);
    const files = folder.getFiles();
    const ss = SpreadsheetApp.openById(currentTrackingSheetId);
    const sheet = ss.getSheetByName("Jobs") || ss.insertSheet("Jobs");

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["File Name", "Job ID", "Status", "Timestamp"]);
      SpreadsheetApp.flush();
    }

    const dataRange = sheet.getDataRange();
    const dataValues = dataRange.getValues();
    const existingDataSet = new Set(dataValues.slice(1).map(row => String(row[0] || "").trim()).filter(Boolean));

    let fileCount = 0;
    let errorCount = 0;
    Logger.log(`[Process] Checking Drive folder "${folder.getName()}" (ID: ${currentFolderId})`);

    while (files.hasNext()) {
      const file = files.next();
      const originalFileName = file.getName();
      const fileId = file.getId();
      const mimeType = file.getMimeType();

      if (!mimeType || !mimeType.startsWith('video/')) continue;
      if (existingDataSet.has(originalFileName.trim())) continue;

      const { studentName, classDate } = extractStudentInfoFromFilename(originalFileName);
      if (!studentName || !classDate) {
        Logger.log(`[Process] Skipping "${originalFileName}": Unable to extract valid student name or date.`);
        continue;
      }

      const standardizedStudentName = studentName.replace(/\s+/g, "_");
      const fileIdPrefix = fileId.substring(0, 10);
      const newFileName = `${standardizedStudentName}_${classDate}_${fileIdPrefix}.mp4`;

      if (existingDataSet.has(newFileName)) {
        Logger.log(`[Process] Skipping "${originalFileName}": Target renamed file "${newFileName}" already exists in sheet.`);
        continue;
      }

      try {
        file.setName(newFileName);
        Logger.log(`[Process] Renamed "${originalFileName}" to "${newFileName}" in Google Drive.`);
      } catch (renameError) {
        Logger.log(`[ERROR] Failed to rename file ID ${fileId}: ${renameError.message}. Skipping.`);
        errorCount++;
        continue;
      }

      sheet.appendRow([newFileName, "", "sending_to_cloudrun", new Date()]);
      const addedRowIndex = sheet.getLastRow();
      SpreadsheetApp.flush();
      existingDataSet.add(newFileName);
      fileCount++;

      Logger.log(`[Process] Sending info to Cloud Run for: ${newFileName}`);
      let finalStatus = "cloudrun_error";
      try {
        const cloudRunSuccess = sendFileInfoToCloudRun(file, newFileName, currentCloudRunUrl);
        if (cloudRunSuccess) {
          finalStatus = "processing_transcription";
        }
      } catch (e) {
        errorCount++;
        Logger.log(`[ERROR] Exception during Cloud Run call for ${newFileName}: ${e.message}`);
      }

      sheet.getRange(addedRowIndex, 3).setValue(finalStatus);
      sheet.getRange(addedRowIndex, 4).setValue(new Date());
      SpreadsheetApp.flush();
      Logger.log(`[Process] Updated sheet for "${newFileName}". Status: ${finalStatus}`);
    }

    Logger.log(`[Process] Finished. Processed ${fileCount} new files. Encountered ${errorCount} errors.`);
  } catch (e) {
    Logger.log(`[Process Error] An unexpected error occurred: ${e.message}\nStack: ${e.stack}`);
  }
}

/**
 * Calls the Cloud Run endpoint to process a file. Confirms receipt and returns true/false.
 */
function sendFileInfoToCloudRun(file, newFileName, currentCloudRunUrl) {
  if (!currentCloudRunUrl) throw new Error("CLOUD_RUN_URL property missing");
  
  const payload = {
    fileId: file.getId(),
    fileName: newFileName
  };
  const options = {
    method: 'POST',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  let response;
  try {
    response = UrlFetchApp.fetch(currentCloudRunUrl, options);
  } catch (e) {
    Logger.log(`[ERROR] Network error calling Cloud Run for ${newFileName}: ${e.message}`);
    throw new Error(`Network-level error calling Cloud Run: ${e.message}`);
  }

  const statusCode = response.getResponseCode();
  const responseText = response.getContentText();

  if (statusCode === 200) {
    Logger.log(`[SUCCESS] Cloud Run call for ${newFileName} returned HTTP 200. Response: ${responseText}`);
    try {
      const result = JSON.parse(responseText);
      if (result && result.message === "File saved to /incoming") {
        return true;
      } else {
        Logger.log(`[WARN] Cloud Run returned 200, but message was unexpected: ${responseText}`);
        throw new Error(`Cloud Run returned an unexpected 200 response.`);
      }
    } catch (parseError) {
      Logger.log(`[ERROR] Failed to parse success response from Cloud Run for ${newFileName}: ${parseError.message}. Response: ${responseText}`);
      throw new Error(`Failed to parse Cloud Run success response.`);
    }
  } else {
    Logger.log(`[ERROR] Cloud Run call for ${newFileName} returned status ${statusCode}: ${responseText}`);
    throw new Error(`Cloud Run returned non-200 status: ${statusCode}`);
  }
}

/**
 * Extracts student name and date (YYYY-MM-DD) from various filename formats.
 */
function extractStudentInfoFromFilename(fileName) {
  let cleanedFileName = fileName.replace(/\.[^/.]+$/, "").trim();
  cleanedFileName = cleanedFileName.replace(/\//g, "-");
  const meetRegex = /^(.+?)[ _-]+(\d{4}-\d{2}-\d{2}).*$/;
  const processedRegex = /^(.+?)_(\d{4}-\d{2}-\d{2})(?:_[A-Za-z0-9_-]{10})?$/;
  let match = cleanedFileName.match(meetRegex) || cleanedFileName.match(processedRegex);

  if (match && match[1] && match[2]) {
    const rawNamePart = match[1];
    const specialCharsRegex = /[~^*'`+]/g;
    const namePartWithoutSpecialChars = rawNamePart.replace(specialCharsRegex, '').trim();
    const studentName = namePartWithoutSpecialChars
      .replace(/_/g, " ")
      .replace(/\s+/g, ' ')
      .replace(/[ -]+$/, "")
      .trim();
    const classDate = match[2];

    if (studentName && /^\d{4}-\d{2}-\d{2}$/.test(classDate)) {
       return { studentName, classDate };
    } else {
       Logger.log(`[WARN extractInfo] Extracted empty name or invalid date from: ${fileName}`);
       return { studentName: null, classDate: null };
    }
  }
  Logger.log(`[ERROR extractInfo] Could not extract student info pattern from: ${fileName}`);
  return { studentName: null, classDate: null };
}

// ------------------------------
// HOMEWORK AND POST-PROCESSING
// ------------------------------

/**
 * [PROPOSED NEW FUNCTION]
 * Scans student Drive folders for new transcripts and triggers homework assignment.
 * This function can be run on a time-based trigger (e.g., every hour).
 * It replaces the old `importTranscriptsToDrive` function.
 */
function processCompletedTranscripts() {
  Logger.log("[Homework] Starting completed transcript processing...");
  const props = PropertiesService.getScriptProperties();
  const rosterId = props.getProperty('STUDENT_ROSTER_ID');
  if (!rosterId) { Logger.log("[Homework] STUDENT_ROSTER_ID not set. Aborting."); return; }

  const homeworkLedgerSheetName = props.getProperty('HOMEWORK_SHEET_NAME');
  const rosterSS = SpreadsheetApp.openById(rosterId);
  const rosterSheet = rosterSS.getSheetByName("Current_Students");
  const homeworkLedgerSheet = rosterSS.getSheetByName(homeworkLedgerSheetName);

  if (!rosterSheet) { Logger.log("[Homework] 'Current_Students' sheet not found."); return; }
  if (!homeworkLedgerSheet) { Logger.log(`[Homework] '${homeworkLedgerSheetName}' sheet not found.`); return; }

  const rosterData = rosterSheet.getDataRange().getValues();
  const homeworkData = homeworkLedgerSheet.getDataRange().getValues();
  // Create a set of already processed transcript filenames from the prompt file names in the homework ledger
  const header = homeworkData[0].map(h => String(h).toLowerCase().trim());
  const promptFileIdCol = header.indexOf('promptfileid');
  const processedTranscripts = new Set();
  if(promptFileIdCol !== -1) {
    for(let i = 1; i < homeworkData.length; i++) {
      const promptFileId = homeworkData[i][promptFileIdCol];
      try {
        const promptFileName = DriveApp.getFileById(promptFileId).getName();
        // Extract transcript name from prompt name, e.g., 'prompt_Student_Name_..._.txt' -> 'Student_Name_..._.txt'
        const transcriptName = promptFileName.replace(/^prompt_/, '');
        processedTranscripts.add(transcriptName);
      } catch (e) {
        Logger.log(`[Homework] Warning: Could not access prompt file with ID ${promptFileId}. It might be deleted.`);
      }
    }
  }
  Logger.log(`[Homework] Found ${processedTranscripts.size} transcripts already processed for homework.`);

  const studentHeader = rosterData[0].map(h => String(h).toLowerCase().trim());
  const nameCol = studentHeader.indexOf("student_name");
  const folderIdCol = studentHeader.indexOf("drive_folder_id");
  const emailCol = studentHeader.indexOf("student_email");

  if ([nameCol, folderIdCol, emailCol].includes(-1)) {
    Logger.log("[Homework] Roster is missing student_name, drive_folder_id, or student_email column.");
    return;
  }
  
  // Group new transcripts by student
  const newTranscriptsByStudent = new Map();

  for (let i = 1; i < rosterData.length; i++) {
    const studentEmail = rosterData[i][emailCol];
    const folderId = rosterData[i][folderIdCol];

    if (!studentEmail || !folderId) continue;
    
    try {
      const studentFolder = DriveApp.getFolderById(folderId);
      const transcriptFiles = studentFolder.getFilesByType(MimeType.PLAIN_TEXT);

      while (transcriptFiles.hasNext()) {
        const transcriptFile = transcriptFiles.next();
        const transcriptName = transcriptFile.getName();

        // Check if this transcript has been processed for homework already
        if (!processedTranscripts.has(transcriptName) && !transcriptName.startsWith('prompt_')) {
          if (!newTranscriptsByStudent.has(studentEmail)) {
            newTranscriptsByStudent.set(studentEmail, []);
          }
          newTranscriptsByStudent.get(studentEmail).push(transcriptName);
        }
      }
    } catch(e) {
      Logger.log(`[Homework] Error processing folder for ${studentEmail} (ID: ${folderId}): ${e.message}`);
    }
  }

  // Assign homework for each student with new transcripts
  Logger.log(`[Homework] Found new transcripts for ${newTranscriptsByStudent.size} students.`);
  for (const [studentEmail, transcriptFileNames] of newTranscriptsByStudent.entries()) {
    try {
      Logger.log(`[Homework] Assigning homework to ${studentEmail} for: ${transcriptFileNames.join(', ')}`);
      assignHomework(studentEmail, transcriptFileNames);
    } catch (e) {
      Logger.log(`[Homework] ERROR assigning homework for ${studentEmail}: ${e.message}`);
    }
  }
  Logger.log("[Homework] Finished processing completed transcripts.");
}


/* ==================================================================== */
/*  EDU SCRIBE  –  HOMEWORK PROMPT UTILITIES                            */
/* ==================================================================== */

/* --- Column indices in ‘Current_Students’ (0-based) --- */
const COL_ROSTER_STUDENT_NAME       = 0;
const COL_ROSTER_STUDENT_ID         = 1;
const COL_ROSTER_LIFE_AND_LIFESTYLE = 16;

/* --- Global IDs pulled once from Script Properties --- */
const SPREAD_ID       = PropertiesService.getScriptProperties().getProperty('STUDENT_ROSTER_ID');
const HW_SHEET_NAME   = PropertiesService.getScriptProperties().getProperty('HOMEWORK_SHEET_NAME');
const DRIVE_COL_PROP  = PropertiesService.getScriptProperties().getProperty('ROSTER_DRIVE_FOLDER_ID');

function rosterFindByEmail(email) {
  if (!email) { Logger.log('[rosterFindByEmail] Called with null/empty email.'); return null; }
  if (!SPREAD_ID) { Logger.log('[rosterFindByEmail] SPREAD_ID is not set.'); return null; }
  
  const sheet = SpreadsheetApp.openById(SPREAD_ID).getSheetByName('Current_Students');
  if (!sheet) { Logger.log(`[rosterFindByEmail] Sheet 'Current_Students' not found.`); return null; }
  
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return null;

  const headerLower = values[0].map(h => String(h).toLowerCase().trim());
  const emailColIdx = headerLower.indexOf('student_email');
  if (emailColIdx === -1) { Logger.log('[rosterFindByEmail] Column "Student_Email" not found.'); return null; }
  
  const driveColIdx = isNaN(DRIVE_COL_PROP) ? -1 : Number(DRIVE_COL_PROP);
  if (driveColIdx === -1) { Logger.log('[rosterFindByEmail] ROSTER_DRIVE_FOLDER_ID is missing/invalid.'); return null; }

  const target = email.toLowerCase().trim();
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const rowEmail = row[emailColIdx];
    if (rowEmail && typeof rowEmail === 'string' && rowEmail.trim().toLowerCase() === target) {
       return {
        Student_ID       : row[COL_ROSTER_STUDENT_ID]?.toString() || '',
        Name             : row[COL_ROSTER_STUDENT_NAME]?.toString() || '',
        Email            : rowEmail.trim(),
        Drive_Folder_ID  : row[driveColIdx]?.toString() || '',
        LifeStyle        : row[COL_ROSTER_LIFE_AND_LIFESTYLE]?.toString() || '',
        RosterRowIndex   : r + 1
      };
    }
  }
  Logger.log(`[rosterFindByEmail] Email "${email}" not found.`);
  return null;
}

function saveHomeworkPrompt(studentEmail, hwId, basePromptText, promptFileName) {
  const student = rosterFindByEmail(studentEmail);
  if (!student) throw new Error(`Student email ${studentEmail} not found.`);
  if (!student.Drive_Folder_ID) throw new Error(`Drive_Folder_ID missing for ${studentEmail}.`);

  const fullPrompt = `${basePromptText.trim()}\n\n---\nStudent background / Life & Lifestyle\n${(student.LifeStyle || 'No profile data on record.').trim()}`;
  let promptFile;
  try {
    const targetFolder = DriveApp.getFolderById(student.Drive_Folder_ID);
    promptFile = targetFolder.createFile(promptFileName, fullPrompt, MimeType.PLAIN_TEXT);
    Logger.log(`[saveHomeworkPrompt] Created prompt file "${promptFileName}" for ${studentEmail}.`);
  } catch (e) {
    throw new Error(`Failed to create prompt file: ${e.message}`);
  }

  const token = Utilities.getUuid();
  try {
    const ledgerSheet = SpreadsheetApp.openById(SPREAD_ID).getSheetByName(HW_SHEET_NAME);
    if (!ledgerSheet) throw new Error(`Ledger sheet "${HW_SHEET_NAME}" not found.`);
    
    const newRowData = [
      student.Student_ID || '', student.Name || '', student.Email || '', hwId || '',
      promptFile.getId(), token, new Date().toISOString(), '', 'Active', 0
    ];
    ledgerSheet.appendRow(newRowData);
    SpreadsheetApp.flush();
    Logger.log(`[saveHomeworkPrompt] Appended to ledger for HW_ID: ${hwId}, Token: ${token}`);
  } catch (e) {
    Logger.log(`[saveHomeworkPrompt] ERROR appending ledger row: ${e.message}`);
    try { if (promptFile) promptFile.setTrashed(true); } catch (trashErr) {}
    throw new Error(`Failed to append ledger row: ${e.message}`);
  }
  return token;
}

function assignHomework(studentEmail, transcriptFileNames) {
  if (!studentEmail || !Array.isArray(transcriptFileNames) || transcriptFileNames.length === 0) {
    throw new Error("Missing parameters for assignHomework.");
  }
  const student = rosterFindByEmail(studentEmail);
  if (!student) throw new Error(`Student lookup failed for ${studentEmail}`);

  const firstTranscriptName = transcriptFileNames[0];
  const transcriptMatch = firstTranscriptName.match(/_(\d{4}-\d{2}-\d{2})_([A-Za-z0-9_-]{10})\.txt$/i);
  let promptFileName = `prompt_${firstTranscriptName}`; // Fallback naming
  if (transcriptMatch?.[1] && transcriptMatch?.[2]) {
    promptFileName = `prompt_${student.Name.replace(/\s+/g, '_')}_${transcriptMatch[1]}_${transcriptMatch[2]}.txt`;
  }

  const todayForHwId = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
  const hwId = `L${student.RosterRowIndex || 'X'}-${todayForHwId}`;
  const transcriptList = transcriptFileNames.join(', ');
  const basePrompt = `You are a GPT called Homework Coach.\nTranscript files for this session: ${transcriptList}.\nThe exact grammar topic being studied might be mentioned directly in the conversation. Search for this to identify what it is, but if more than one grammar topic or grammar theme is mentioned, then surmize the grammar topic being taught, judging from repeated sentence structures being practiced and especially how the teacher introduces the topic and corrects the student. Carefully analyze the student’s communications and reactions to the teacher’s instructions to find instances in which the grammar being taught was not well comprehended by the student and based on this analysis, choosing only one grammatical theme to pursue and formulate all ten questions and dialogue around this one theme for this homework session. In your greeting, explicitly state the grammar topic being covered in the homework, followed by a very brief situation based on the student's life and lifestyle data or a relevant comment from the student transcript to provide context showing how the grammar structure can be productively applied. Ask up to 10 personalized questions based on the conclusions you reached in your analysis of the transcript and incorporate the student background provided below. Show the problem/question number (i) for each question asked. After student answers, provide an instant, short and empathetic evaluation of the student's answer, explaining clearly but briefly why the student is right or what needs work, then in the same output, move on to the next student prompt, with the next number (i+1) displayed for reference. Focus on reinforcing weak points with this grammar topic identified in the transcript(s), being sure to add professional and personal details from the transcript and the student information below to keep the output relevant to the student’s profession, personal characteristics and world view. General behavior when interacting with the student: *Do not prompt the student to SPEAK or LISTEN to you. You will be interacting by text chat only so speaking and listening is not possible. *Be friendly and empathetic but brief in your responses. *When you’ve finished coaching the student through the 10 homework problems, it is essential to bring the conversation to a warm and polite close by instructing the student to click the button that says ‘Done ✅ Submit Homework’.`;
  const token = saveHomeworkPrompt(studentEmail, hwId, basePrompt, promptFileName);

  const portalBaseUrl = PropertiesService.getScriptProperties().getProperty('HOMEWORK_PORTAL_BASEURL');
  if (!portalBaseUrl) {
    Logger.log("[assignHomework] WARNING: HOMEWORK_PORTAL_BASEURL property not set. Skipping email.");
    return;
  }
 
  const portalLink = `${portalBaseUrl}/homework-coach/?token=${token}`;
  const studentFirstName = student.Name ? student.Name.split(' ')[0] : "Student";

  try {
    MailApp.sendEmail(studentEmail, `Your new homework is ready (ID: ${hwId})`, `Hi ${studentFirstName},\n\nYour practice set based on your last class is ready. Click the link below to open it in your portal:\n\n${portalLink}\n\nRemember: complete all ten turns with the Homework GPT, then hit Done ✅ to turn in your homework.\n\nGood luck!\n`);
    Logger.log(`[assignHomework] HW ${hwId} email sent to ${studentEmail}`);
  } catch (e) {
    Logger.log(`[assignHomework] ERROR sending email to ${studentEmail} for HW ${hwId}: ${e.message}`);
  }
}

// ------------------------------
// TESTING FUNCTIONS
// ------------------------------

// NOTE: Some of these tests are for the old Speechmatics workflow and are now deprecated.

function testAuthToken() { 
  try {
    const token = getServiceAccountToken();
    Logger.log(`SUCCESS: Token retrieved: ${token.substring(0, 30)}...`);
  } catch (e) {
    Logger.log(`ERROR: ${e.message}`);
  }
}

function testFileAccess() { 
  const props = PropertiesService.getScriptProperties();
  const folderId = props.getProperty("FOLDER_ID");
  if (!folderId) { Logger.log("ERROR: FOLDER_ID not set."); return; }
  try {
    const folder = DriveApp.getFolderById(folderId);
    Logger.log(`Successfully accessed folder: "${folder.getName()}".`);
  } catch (e) {
    Logger.log(`Error accessing Drive folder ID ${folderId}: ${e.message}`);
  }
}

function testStudentRosterAccess() { 
  const props = PropertiesService.getScriptProperties();
  const rosterId = props.getProperty("STUDENT_ROSTER_ID");
  if (!rosterId) { Logger.log("ERROR: STUDENT_ROSTER_ID not set."); return; }
  try {
    const ss = SpreadsheetApp.openById(rosterId);
    const sheet = ss.getSheetByName("Current_Students");
    if (!sheet) { Logger.log("ERROR: 'Current_Students' sheet not found."); return; }
    Logger.log(`Successfully accessed roster: "${ss.getName()}".`);
  } catch (e) {
    Logger.log(`Error accessing student roster (ID: ${rosterId}): ${e.message}`);
  }
}

function checkCloudRunUrl() { 
  const cloudRunUrl = PropertiesService.getScriptProperties().getProperty("CLOUD_RUN_URL");
  if (cloudRunUrl) {
    Logger.log(`Cloud Run URL is set to: ${cloudRunUrl}`);
  } else {
    Logger.log("ERROR: CLOUD_RUN_URL is not set in Script Properties.");
  }
}

function testAssignHomework() {
  const testEmail = 'teacher@fakeemail.com'; // Replace with a real test email
  const testTranscriptFileNames = [ `Test_Student_2024-07-01_abcdef1234.txt` ];
  if (testEmail.startsWith('replace.me')) { Logger.log("ERROR: Please replace with a valid email first."); return; }
  try {
    assignHomework(testEmail, testTranscriptFileNames);
    Logger.log(`[SUCCESS] testAssignHomework executed for ${testEmail}.`);
  } catch (e) { 
    Logger.log(`[ERROR] Test failed: ${e.message}\nStack: ${e.stack || 'N/A'}`); 
  }
}