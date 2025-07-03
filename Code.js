/**
 * ========== COMPLETE UPDATED SCRIPT V8 - Corrected Portal Link & Turns_Used Init ==========
 * Google Meet to Speechmatics Integration via Cloud Run
 * Apps Script handles Drive monitoring, renaming, calling Cloud Run,
 * job tracking, monitoring, completion handling, transcript import,
 * and **consolidated homework assignment per student per day.**
 * Includes original testing functions and homework utility functions.
 */


// --- Ensure OAuth2 Library is added: ID 1B7FSrk5Zi6L1rSxxTDgDEUsPzlukDsi4KGuTMorsTQHhGBzBkMun4iDF ---
// Note: OAuth2 library might not be strictly needed anymore if only getServiceAccountToken is used for GCS,
// but doesn't hurt to leave if other potential uses exist. getServiceAccountToken uses manual JWT signing.


// ------------------------------
// GLOBAL CONFIG (LOADED FROM PropertiesService)
// ------------------------------
// Script Properties Keys Needed:
// FOLDER_ID                  (Drive folder for new Meet recordings)
// BUCKET_NAME                (GCS bucket name)
// TRACKING_SHEET_ID          (ID of Spreadsheet for Speechmatics job tracking)
// CLOUD_RUN_URL              (URL of the Cloud Run service for Speechmatics upload/submission)
// SPEECHMATICS_API_KEY       (Your Speechmatics API key)
// STUDENT_ROSTER_ID          (ID of Spreadsheet containing 'Current_Students' and 'Homework_Push' tabs)
// CLIENT_EMAIL               (Service Account email)
// PRIVATE_KEY                (Service Account private key)
// HOMEWORK_SHEET_NAME        (Name of the homework ledger tab, e.g., "Homework_Push")
// ROSTER_DRIVE_FOLDER_ID     (0-based index of the Drive_Folder_ID column in 'Current_Students')
// HOMEWORK_PORTAL_BASEURL    (Base URL for the homework portal link, e.g., https://english-gpt.chat)


/**
 * Retrieves a short-lived OAuth2 access token for the service account by manually signing a JWT.
 * Requires CLIENT_EMAIL and PRIVATE_KEY properties to be set in Script Properties.
 * Used for GCS operations performed directly by Apps Script (monitoring, handling, import).
 * @return {string} The access token.
 * @throws {Error} If authentication fails or properties are missing.
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
    let signatureBytes;
    try {
      signatureBytes = Utilities.computeRsaSha256Signature(signatureInput, privateKey);
    } catch (signError) {
      Logger.log(`[AUTH ERROR] Failed compute RSA signature. Check PRIVATE_KEY format/validity. Error: ${signError.message}`);
      throw new Error(`Failed compute RSA signature. Check PRIVATE_KEY format/validity. Error: ${signError.message}`);
    }
    const encodedSignature = Utilities.base64EncodeWebSafe(signatureBytes);
    const assertion = `${signatureInput}.${encodedSignature}`;
    const tokenResponse = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
      method: 'post',
      contentType: 'application/x-www-form-urlencoded',
      payload: {
        grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
        assertion: assertion
      },
      muteHttpExceptions: true
    });


    const responseCode = tokenResponse.getResponseCode();
    const responseText = tokenResponse.getContentText();


    if (responseCode !== 200) {
      Logger.log(`[AUTH ERROR] Google token endpoint failed (${responseCode}). Response: ${responseText}`);
      let googleErrorDetail = "";
      try {
        const errorJson = JSON.parse(responseText);
        googleErrorDetail = errorJson.error_description || errorJson.error || "";
      } catch (parseError) {}
      throw new Error(`Failed to fetch token from Google (${responseCode}). ${googleErrorDetail || 'Check service account permissions/status.'}`);
    }
    const tokenData = JSON.parse(responseText);
    if (!tokenData.access_token) {
      Logger.log(`[AUTH ERROR] Access token missing in Google's 200 OK response: ${responseText}`);
      throw new Error("Access token not found in Google's response.");
    }
    return tokenData.access_token;
  } catch (e) {
    Logger.log(`[AUTH ERROR] Unexpected exception in getServiceAccountToken: ${e.message}\nStack: ${e.stack}`);
    throw new Error(`[AUTH ERROR] Authentication failed: ${e.message}`);
  }
}




// ------------------------------
// PROCESS NEW RECORDINGS (Trigger Function)
// ------------------------------


/**
 * Finds new video recordings in Drive, renames them, calls Cloud Run
 * to upload and submit to Speechmatics, and updates the tracking sheet.
 */
function processNewRecordings() {
  // Fetch properties needed just for this function run
  const props = PropertiesService.getScriptProperties();
  const currentFolderId = props.getProperty("FOLDER_ID");
  const currentTrackingSheetId = props.getProperty("TRACKING_SHEET_ID");
  const currentCloudRunUrl = props.getProperty("CLOUD_RUN_URL");


  // Validate essential properties
  if (!currentFolderId) { Logger.log("ERROR: FOLDER_ID script property is not set."); return; }
  if (!currentTrackingSheetId) { Logger.log("ERROR: TRACKING_SHEET_ID script property is not set."); return; }
  if (!currentCloudRunUrl) { Logger.log("ERROR: CLOUD_RUN_URL script property is not set."); return; }


  try {
    const folder = DriveApp.getFolderById(currentFolderId);
    const files = folder.getFiles();
    const ss = SpreadsheetApp.openById(currentTrackingSheetId);
    const sheet = ss.getSheetByName("Jobs") || ss.insertSheet("Jobs"); // Use "Jobs" sheet


    // Ensure header row exists
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["File Name", "Job ID", "Status", "Timestamp"]); // Added Timestamp
      SpreadsheetApp.flush();
    } else if (sheet.getRange(1, 1).getValue() !== "File Name") {
      Logger.log("WARNING: Header row in 'Jobs' sheet seems incorrect. Re-writing.");
      sheet.getRange("1:1").clearContent();
      sheet.getRange("A1:D1").setValues([["File Name", "Job ID", "Status", "Timestamp"]]);
    }


    // Get existing filenames to avoid duplicates
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


      if (!mimeType || !mimeType.startsWith('video/')) {
        continue;
      }


      const alreadyRenamedRegex = /^(.+?)_(\d{4}-\d{2}-\d{2})_[A-Za-z0-9_-]{10}\.mp4$/i;
      let isAlreadyTracked = existingDataSet.has(originalFileName.trim());
      let potentialRenamedName = "";


      if (!isAlreadyTracked) {
        const renameMatch = originalFileName.match(alreadyRenamedRegex);
        if (renameMatch) {
          potentialRenamedName = originalFileName.trim();
          isAlreadyTracked = existingDataSet.has(potentialRenamedName);
        }
      }


      if (isAlreadyTracked) {
        continue;
      }


      const { studentName, classDate } = extractStudentInfoFromFilename(originalFileName);


      if (!studentName || !classDate) {
        Logger.log(`[Process] Skipping "${originalFileName}": Unable to extract valid student name or date.`);
        continue;
      }


      const standardizedStudentName = studentName.replace(/\s+/g, "_");
      const fileIdPrefix = fileId.substring(0, 10);
      const newFileName = `${standardizedStudentName}_${classDate}_${fileIdPrefix}.mp4`;


      if (existingDataSet.has(newFileName)) {
        Logger.log(`[Process] Skipping "${originalFileName}": Target renamed file "${newFileName}" already exists in tracking sheet.`);
        continue;
      }


      try {
        file.setName(newFileName);
        Logger.log(`[Process] Renamed "${originalFileName}" to "${newFileName}" in Google Drive.`);
      } catch (renameError) {
        Logger.log(`[ERROR] Failed to rename file ID ${fileId} from "${originalFileName}" to "${newFileName}". Error: ${renameError.message}. Skipping this file.`);
        errorCount++;
        continue;
      }


      const timestamp = new Date();
      sheet.appendRow([newFileName, "", "processing_cloudrun", timestamp]);
      const addedRowIndex = sheet.getLastRow();
      SpreadsheetApp.flush();
      existingDataSet.add(newFileName);
      fileCount++;


      Logger.log(`[Process] Sending renamed file info to Cloud Run: ${newFileName}`);
      let jobId = null;
      let cloudRunError = false;
      try {
        jobId = sendFileInfoToCloudRun(file, newFileName, currentCloudRunUrl);
      } catch (e) {
        cloudRunError = true;
        errorCount++;
        // Log the detailed error from sendFileInfoToCloudRun if needed
        Logger.log(`[ERROR] Exception during Cloud Run call for ${newFileName}: ${e.message}`);
      }


      let finalStatus = "pending_submission_check"; // Default if no error and no jobId
      if (cloudRunError) {
        finalStatus = "cloudrun_error";
      } else if (jobId) {
        finalStatus = "submitted";
      } else {
        Logger.log(`[WARN] Cloud Run call for "${newFileName}" did not return a valid Job ID or encountered an error.`);
        finalStatus = "cloudrun_no_jobid";
      }


      sheet.getRange(addedRowIndex, 2).setValue(jobId || "");
      sheet.getRange(addedRowIndex, 3).setValue(finalStatus);
      sheet.getRange(addedRowIndex, 4).setValue(new Date()); // Update timestamp to reflect processing attempt
      SpreadsheetApp.flush();
      Logger.log(`[Process] Updated sheet row ${addedRowIndex} for "${newFileName}". JobID: ${jobId || 'N/A'}, Status: ${finalStatus}`);


    } // End while loop


    Logger.log(`[Process] Finished checking Drive folder. Processed ${fileCount} new files. Encountered ${errorCount} errors during rename or Cloud Run call initiation.`);


  } catch (e) {
    Logger.log(`[Process Error] An unexpected error occurred in processNewRecordings: ${e.message}`);
    Logger.log(`[Process Error] Stack: ${e.stack}`);
  }
}


/**
 * Calls the Cloud Run endpoint to process a file.
 * Parses the response to extract the job ID.
 */
function sendFileInfoToCloudRun(file, newFileName, currentCloudRunUrl) {
  if (!currentCloudRunUrl) {
    throw new Error("CLOUD_RUN_URL property missing");
  }


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
    Logger.log(`[ERROR] Network or UrlFetchApp error calling Cloud Run for ${newFileName}: ${e.message}`);
    if (e.message.toLowerCase().includes('timeout')) {
       Logger.log(`[Hint] This might be a timeout error. Consider async processing.`);
    }
    return null;
  }


  const statusCode = response.getResponseCode();
  const responseText = response.getContentText();


  if (statusCode !== 200) {
    Logger.log(`[ERROR] Cloud Run call for ${newFileName} returned status ${statusCode}: ${responseText}`);
    try {
        const errorJson = JSON.parse(responseText);
        Logger.log(`[Cloud Run Error Detail]: ${JSON.stringify(errorJson)}`);
    } catch (parseError) { /* Ignore */ }
    return null;
  } else {
    Logger.log(`[SUCCESS] Cloud Run response for ${newFileName} (HTTP 200): ${responseText}`);
    try {
      const result = JSON.parse(responseText);
      if (result && result.job_id && typeof result.job_id === 'string' && result.job_id.trim() !== "") {
        return result.job_id.trim();
      } else {
        Logger.log(`[WARN] Cloud Run status 200 for ${newFileName}, but 'job_id' missing or invalid in response: ${responseText}`);
        return null;
      }
    } catch (parseError) {
      Logger.log(`[ERROR] Failed to parse successful JSON response from Cloud Run for ${newFileName}: ${parseError.message}. Response: ${responseText}`);
      return null;
    }
  }
}




// ------------------------------
// EXTRACT STUDENT INFO FROM FILENAME
// ------------------------------
/**
 * Extracts student name and date (YYYY-MM-DD) from various filename formats.
 * Handles Google Meet format and the standardized format used after renaming.
 * Cleans potential problematic characters from the name part.
 * @param {string} fileName The original or renamed filename.
 * @return {{studentName: string|null, classDate: string|null}} Object containing extracted info.
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
       // Logger.log(`[extractInfo] Extracted: Name='${studentName}', Date='${classDate}' from "${fileName}"`); // Optional: verbose log
       return { studentName, classDate };
    } else {
       Logger.log(`[WARN extractInfo] Extracted empty name or invalid date after cleaning from: ${fileName} (Name: '${studentName}', Date: '${classDate}')`);
       return { studentName: null, classDate: null };
    }
  }
  Logger.log(`[ERROR extractInfo] Could not extract student info pattern from: ${fileName}`);
  return { studentName: null, classDate: null };
}




// ---------------------------------------------------------------------------------
// --- SUBMIT TRANSCRIPTIONS TO SPEECHMATICS (REMOVED - Handled by Cloud Run) ---
// ---------------------------------------------------------------------------------
/* Functions submitNewTranscriptions() and submitToSpeechmatics() are removed. */




// ------------------------------
// MONITOR SPEECHMATICS JOBS
// ------------------------------
/**
 * Checks the status of submitted Speechmatics jobs and processes completed ones.
 */
function monitorTranscriptionJobs() {
  const props = PropertiesService.getScriptProperties();
  const currentTrackingSheetId = props.getProperty("TRACKING_SHEET_ID");
  const currentBucketName = props.getProperty("BUCKET_NAME");
  const currentSpeechmaticsApiKey = props.getProperty("SPEECHMATICS_API_KEY");


  if (!currentTrackingSheetId) { Logger.log("ERROR: TRACKING_SHEET_ID script property is not set."); return; }
  if (!currentBucketName) { Logger.log("ERROR: BUCKET_NAME script property is not set."); return; }
  if (!currentSpeechmaticsApiKey) { Logger.log("ERROR: SPEECHMATICS_API_KEY script property is not set."); return; }


  Logger.log("[Monitor] Starting monitorTranscriptionJobs...");
  let accessToken;
  try { accessToken = getServiceAccountToken(); }
  catch (authError) { Logger.log(`[ERROR] Auth failed in monitorTranscriptionJobs: ${authError.message}. Aborting monitor.`); return; }


  const ss = SpreadsheetApp.openById(currentTrackingSheetId);
  const sheet = ss.getSheetByName("Jobs");
  if (!sheet || sheet.getLastRow() < 2) { Logger.log("[Monitor] 'Jobs' sheet not found or empty."); return; }


  const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3);
  const data = dataRange.getValues();
  let checkedCount = 0, completedCount = 0, errorCount = 0;
  Logger.log(`[Monitor] Checking ${data.length} rows.`);


  const terminalStatuses = new Set([
    "done", "rejected", "processing_error", "gcs_error", "transcript_fetch_error",
    "submission_error", "submission_exception", "cloudrun_error", "cloudrun_no_jobid",
    "not_found", "error_fetching_status", "exception_fetching_status", "error_parsing_response",
    "error_unauthorized", "error_missing_jobid", "gcs_move_error"
  ]);
  const submittedStatuses = new Set(["submitted", "running", "processing"]);


  for (let i = 0; i < data.length; i++) {
    const fileName = String(data[i][0] || "").trim();
    const jobId = String(data[i][1] || "").trim();
    let localStatus = String(data[i][2] || "").trim().toLowerCase();
    const currentRow = i + 2;
    const statusCell = sheet.getRange(currentRow, 3);


    if (!fileName || terminalStatuses.has(localStatus)) continue;


    if ((submittedStatuses.has(localStatus) || localStatus === "processing_transcript") && !jobId) {
       if (localStatus !== "error_missing_jobid") {
           Logger.log(`[Monitor] Row ${currentRow}: File "${fileName}" status "${localStatus}" but no Job ID. Marking error.`);
           statusCell.setValue("error_missing_jobid"); SpreadsheetApp.flush(); errorCount++;
       }
       continue;
    }
    if (!submittedStatuses.has(localStatus) && localStatus !== "processing_transcript") continue;


    checkedCount++;
    let newStatus = "";
    try {
      // Logger.log(`[Monitor DEBUG] Checking status for Job ID: ${jobId} (File: ${fileName})`); // Optional verbose log
      newStatus = getJobStatus(jobId, currentSpeechmaticsApiKey);
      // Logger.log(`[Monitor DEBUG] API returned status: ${newStatus} for Job ID: ${jobId}`); // Optional verbose log
    } catch (e) { Logger.log(`[ERROR] Exception calling getJobStatus for Job ID ${jobId} (File: ${fileName}): ${e.message}`); newStatus = "exception_fetching_status"; }


    const newStatusLower = newStatus.toLowerCase();


    if (newStatusLower === "done") {
      if (localStatus === "done") continue;
      Logger.log(`[Monitor] Job ${jobId} for "${fileName}" complete. Processing...`);
      statusCell.setValue("processing_transcript"); SpreadsheetApp.flush();
      let processingError = false; let specificErrorStatus = "processing_error";
      try {
        handleCompletedJob(fileName, jobId, currentBucketName, accessToken, currentSpeechmaticsApiKey);
        moveFileInGCS(currentBucketName, `incoming/${fileName}`, `processed/${fileName}`, accessToken);
      } catch (err) {
        processingError = true;
        if (err.message.toLowerCase().includes("transcript")) specificErrorStatus = "transcript_fetch_error";
        else if (err.message.toLowerCase().includes("gcs upload")) specificErrorStatus = "gcs_error";
        else if (err.message.toLowerCase().includes("movegcs")) specificErrorStatus = "gcs_move_error";
        Logger.log(`[ERROR] Post-processing failed for "${fileName}" (Job ${jobId}). Status: ${specificErrorStatus}. Error: ${err.message}\nStack: ${err.stack || 'N/A'}`);
      }
      const finalStatus = processingError ? specificErrorStatus : "done";
      Logger.log(`[Monitor] Setting final status for "${fileName}" to "${finalStatus}".`);
      statusCell.setValue(finalStatus); SpreadsheetApp.flush();
      if (!processingError) completedCount++; else errorCount++;
      continue;
    }
    else if (newStatusLower === "rate_limited") { Logger.log(`[Monitor] Rate limited checking Job ID ${jobId}.`); continue; }
    else if (terminalStatuses.has(newStatusLower)) {
        if (newStatusLower !== localStatus) {
             Logger.log(`[Monitor] Updating status for "${fileName}" (Job ${jobId}) from "${localStatus}" to terminal state "${newStatus}".`);
             statusCell.setValue(newStatus); SpreadsheetApp.flush();
             if (newStatusLower.includes("error") || newStatusLower.includes("reject") || newStatusLower === "not_found") errorCount++;
        }
        continue;
    }
    else if (newStatusLower !== localStatus && newStatus !== "") {
      Logger.log(`[Monitor] Updating status for "${fileName}" (Job ${jobId}) from "${localStatus}" to "${newStatus}".`);
      statusCell.setValue(newStatus); SpreadsheetApp.flush();
    }
  } // End for loop
  Logger.log(`[Monitor] Finished. Checked ${checkedCount} active jobs. Completed ${completedCount}. Errors: ${errorCount}.`);
}


/**
 * Fetches the status of a Speechmatics job. Includes refined error handling.
 */
function getJobStatus(jobId, speechmaticsApiKey) {
  if (!jobId || !speechmaticsApiKey) { Logger.log("[ERROR getJobStatus] Missing Job ID or API Key."); return "error_missing_params"; }
  const url = `https://asr.api.speechmatics.com/v2/jobs/${jobId}`;
  try {
    const response = UrlFetchApp.fetch(url, { method: 'GET', headers: { Authorization: `Bearer ${speechmaticsApiKey}` }, muteHttpExceptions: true });
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    if (responseCode === 200) {
      const jsonResponse = JSON.parse(responseText);
      if (jsonResponse?.job?.status) { return jsonResponse.job.status; }
      else { Logger.log(`[ERROR getJobStatus] Job ID ${jobId}: Valid 200 but unexpected JSON: ${responseText}`); return "error_parsing_response"; }
    } else if (responseCode === 401) { Logger.log(`[ERROR getJobStatus] Job ID ${jobId}: Unauthorized (401).`); return "error_unauthorized"; }
    else if (responseCode === 404) { Logger.log(`[WARN getJobStatus] Job ID ${jobId}: Not Found (404).`); return "not_found"; }
    else if (responseCode === 429) { Logger.log(`[WARN getJobStatus] Job ID ${jobId}: Rate limited (429).`); return "rate_limited"; }
    else { Logger.log(`[ERROR getJobStatus] Job ID ${jobId}: HTTP ${responseCode}. ${responseText}`); let smError = responseText; try { smError = JSON.parse(responseText)?.error || responseText; } catch (ignore) {} return `error_fetching_${responseCode}`; }
  } catch (e) { Logger.log(`[ERROR getJobStatus] Job ID ${jobId}: Exception: ${e.message}`); if (e instanceof SyntaxError) return "error_parsing_response"; return "exception_fetching_status"; }
}




// ------------------------------
// HANDLE COMPLETED JOBS
// ------------------------------
/**
 * Handles a completed Speechmatics job: fetches the transcript, uploads it to GCS 'transcripts/' folder.
 * Uses tracking title if available, otherwise original filename.
 */
function handleCompletedJob(fileName, jobId, bucketName, accessToken, apiKey) {
  if (!fileName || !jobId || !bucketName || !accessToken || !apiKey) throw new Error("Missing parameters for handleCompletedJob");
  let trackingTitle = null; const jobDetailsUrl = `https://asr.api.speechmatics.com/v2/jobs/${jobId}`;
  try {
    const jobResponse = UrlFetchApp.fetch(jobDetailsUrl, { headers: { Authorization: `Bearer ${apiKey}` }, muteHttpExceptions: true });
    if (jobResponse.getResponseCode() === 200) { trackingTitle = JSON.parse(jobResponse.getContentText())?.job?.tracking?.title || null; Logger.log(`[HandleComplete DEBUG] Tracking title Job ${jobId}: ${trackingTitle || 'Not found'}`); }
    else { Logger.log(`[WARN HandleComplete] Fetch job details ${jobId} failed (HTTP ${jobResponse.getResponseCode()}).`); }
  } catch (e) { Logger.log(`[WARN HandleComplete] Exc fetching job details ${jobId}: ${e.message}.`); }


  const nameFormatRegex = /.+_\d{4}-\d{2}-\d{2}_[A-Za-z0-9_-]{10}\.mp4$/i;
  const transcriptFileNameBase = (trackingTitle && nameFormatRegex.test(trackingTitle)) ? trackingTitle.replace(/\.mp4$/i, '') : fileName.replace(/\.mp4$/i, '');
  const transcriptFileName = `${transcriptFileNameBase}.txt`;
  Logger.log(`[HandleComplete] Determined transcript filename: "${transcriptFileName}" for Job ${jobId}`);


  let transcriptContent = null; const transcriptUrl = `https://asr.api.speechmatics.com/v2/jobs/${jobId}/transcript?format=txt`;
  try {
    Logger.log(`[HandleComplete] Fetching transcript: ${transcriptUrl}`);
    const transcriptResponse = UrlFetchApp.fetch(transcriptUrl, { headers: { Authorization: `Bearer ${apiKey}` }, muteHttpExceptions: true });
    const transcriptResponseCode = transcriptResponse.getResponseCode(); const responseText = transcriptResponse.getContentText();
    if (transcriptResponseCode !== 200) { Logger.log(`[ERROR HandleComplete] Failed get transcript Job ${jobId}. HTTP ${transcriptResponseCode}: ${responseText}`); let smError = responseText; try { smError = JSON.parse(responseText)?.error || responseText; } catch (ignore) {} throw new Error(`Transcript fetch failed (${transcriptResponseCode}): ${smError}`); }
    transcriptContent = responseText;
    if (transcriptContent === null || transcriptContent.trim() === "") { Logger.log(`[WARN HandleComplete] Job ${jobId}: Transcript empty.`); }
    Logger.log(`[HandleComplete] Retrieved transcript (${transcriptContent?.length || 0} chars) Job ${jobId}.`);
  } catch (e) { Logger.log(`[ERROR HandleComplete] Exc transcript fetch Job ${jobId}: ${e.message}`); throw new Error(`Transcript fetch error: ${e.message}`); }


  const gcsObjectName = `transcripts/${transcriptFileName}`; const uploadUrl = `https://storage.googleapis.com/upload/storage/v1/b/${bucketName}/o?uploadType=media&name=${encodeURIComponent(gcsObjectName)}`;
  try {
    const blob = Utilities.newBlob(transcriptContent || "", 'text/plain', transcriptFileName);
    const uploadOptions = { method: 'POST', contentType: 'text/plain; charset=utf-8', payload: blob.getBytes(), headers: { Authorization: `Bearer ${accessToken}` }, muteHttpExceptions: true };
    Logger.log(`[HandleComplete] Uploading transcript to GCS: ${gcsObjectName}`);
    const uploadResponse = UrlFetchApp.fetch(uploadUrl, uploadOptions);
    const uploadResponseCode = uploadResponse.getResponseCode(); const uploadResponseText = uploadResponse.getContentText();
    if (uploadResponseCode !== 200) { Logger.log(`[ERROR HandleComplete] Failed GCS upload ${gcsObjectName}. HTTP ${uploadResponseCode}: ${uploadResponseText}`); throw new Error(`GCS upload failed (${uploadResponseCode}): ${uploadResponseText}`); }
    Logger.log(`[HandleComplete] Successfully uploaded transcript GCS: ${gcsObjectName}`);
  } catch (e) { Logger.log(`[ERROR HandleComplete] Exc GCS upload ${gcsObjectName}: ${e.message}`); throw new Error(`GCS upload error: ${e.message}`); }
}




// ------------------------------
// MOVE FILE IN GCS (Helper Function)
// ------------------------------
/**
 * Moves a file within a GCS bucket by copying and then deleting the source.
 * Includes check if source exists. Throws specific errors.
 */
function moveFileInGCS(bucketName, sourceObjectName, destinationObjectName, token) {
  if (!bucketName || !sourceObjectName || !destinationObjectName || !token) throw new Error("Missing parameters for moveFileInGCS");
  const baseUrl = `https://storage.googleapis.com/storage/v1/b/${bucketName}/o/`;
  const sourceUrl = baseUrl + encodeURIComponent(sourceObjectName);
  const destinationUrlEncoded = encodeURIComponent(destinationObjectName);
  const checkErrorMsg = `MoveGCS check source error: ${sourceObjectName}`;
  const copyErrorMsg = `MoveGCS copy error: ${sourceObjectName} to ${destinationObjectName}`;
  const deleteErrorMsg = `MoveGCS delete source error: ${sourceObjectName}`;


  try { // Check source
    const headResp = UrlFetchApp.fetch(sourceUrl + '?fields=name', { method: 'GET', headers: { Authorization: `Bearer ${token}` }, muteHttpExceptions: true });
    const headRespCode = headResp.getResponseCode();
    if (headRespCode === 404) { Logger.log(`[WARN MoveGCS] Source not found: ${sourceObjectName}. Skipping.`); return; }
    else if (headRespCode !== 200) { Logger.log(`[ERROR MoveGCS Check] HTTP ${headRespCode}: ${headResp.getContentText()}`); throw new Error(`Source check failed (HTTP ${headRespCode})`); }
  } catch(e) { Logger.log(`[ERROR MoveGCS] ${checkErrorMsg} - ${e.message}`); throw new Error(checkErrorMsg); }


  try { // Copy
    const copyUrl = `${sourceUrl}/copyTo/b/${bucketName}/o/${destinationUrlEncoded}`;
    const copyResp = UrlFetchApp.fetch(copyUrl, { method: 'POST', headers: { Authorization: `Bearer ${token}` }, muteHttpExceptions: true });
    const copyRespCode = copyResp.getResponseCode();
    if (copyRespCode !== 200) { Logger.log(`[ERROR MoveGCS Copy] HTTP ${copyRespCode}: ${copyResp.getContentText()}`); throw new Error(`Copy failed (HTTP ${copyRespCode})`); }
    Logger.log(`[MoveGCS] Copied successfully: ${sourceObjectName} -> ${destinationObjectName}`);
  } catch (e) { Logger.log(`[ERROR MoveGCS] ${copyErrorMsg} - ${e.message}`); throw new Error(copyErrorMsg); }


  try { // Delete
    const deleteResp = UrlFetchApp.fetch(sourceUrl, { method: 'DELETE', headers: { Authorization: `Bearer ${token}` }, muteHttpExceptions: true });
    const deleteRespCode = deleteResp.getResponseCode();
    if (deleteRespCode !== 204) { Logger.log(`[ERROR MoveGCS Delete] HTTP ${deleteRespCode}: ${deleteResp.getContentText()}`); throw new Error(`Delete failed (HTTP ${deleteRespCode})`); }
    Logger.log(`[MoveGCS] Deleted source successfully: ${sourceObjectName}.`);
  } catch (e) { Logger.log(`[ERROR MoveGCS] ${deleteErrorMsg} - ${e.message}`); throw new Error(deleteErrorMsg); }
}




// ------------------------------
// FETCH GCS FILE CONTENT (Helper Function)
// ------------------------------
/** Fetches the content of a text file from GCS. Returns null on error. */
function fetchGCSFileContent(bucketName, objectName, accessToken) {
  if (!bucketName || !objectName || !accessToken) { Logger.log("[ERROR FetchGCS] Missing parameters."); return null; }
  const url = `https://storage.googleapis.com/storage/v1/b/${bucketName}/o/${encodeURIComponent(objectName)}?alt=media`;
  try {
    const response = UrlFetchApp.fetch(url, { headers: { Authorization: `Bearer ${accessToken}` }, muteHttpExceptions: true });
    const responseCode = response.getResponseCode();
    if (responseCode === 200) return response.getContentText();
    else {
      if (responseCode === 404) Logger.log(`[ERROR FetchGCS] ${objectName}: Not Found (404).`);
      else if (responseCode === 403) Logger.log(`[ERROR FetchGCS] ${objectName}: Forbidden (403). Check SA permissions.`);
      else Logger.log(`[ERROR FetchGCS] ${objectName}. HTTP ${responseCode}: ${response.getContentText()}`);
      return null;
    }
  } catch (e) { Logger.log(`[ERROR FetchGCS] Exc fetch ${objectName}: ${e.message}\nStack: ${e.stack}`); return null; }
}




// ------------------------------
// IMPORT TRANSCRIPTS TO DRIVE
// ------------------------------
/**
 * Imports completed transcripts from GCS 'transcripts/' folder to student Google Drive folders.
 * V6: Groups same-day imports per student and calls assignHomework ONCE per student/day after processing all files.
 */
function importTranscriptsToDrive() {
  // Fetch properties needed
  const props = PropertiesService.getScriptProperties();
  const currentStudentRosterId = props.getProperty("STUDENT_ROSTER_ID");
  const currentBucketName = props.getProperty("BUCKET_NAME");
  // const localClientEmail = props.getProperty('CLIENT_EMAIL'); // This variable was unused in this function's scope


  // Validate essential properties
  if (!currentStudentRosterId) { Logger.log("ERROR: STUDENT_ROSTER_ID script property is not set."); return; }
  if (!currentBucketName) { Logger.log("ERROR: BUCKET_NAME script property is not set."); return; }


  Logger.log("[Import V6] Starting transcript import and homework assignment...");
  let accessToken;
  try {
    accessToken = getServiceAccountToken();
  } catch (authError) {
    Logger.log(`[ERROR Import] Auth failed: ${authError.message}. Aborting import.`);
    return;
  }


  // --- Read Roster and Create Map (INCLUDING EMAIL) ---
  let rosterMap; // Map: lowercase processed name -> { folderId, originalName, email }
  try {
    const rosterSS = SpreadsheetApp.openById(currentStudentRosterId);
    const rosterSheet = rosterSS.getSheetByName("Current_Students");
    if (!rosterSheet) { throw new Error("Sheet 'Current_Students' not found."); }
    const rosterData = rosterSheet.getDataRange().getValues();
    rosterMap = new Map();


    if (rosterData.length < 2) {
      Logger.log("[WARN Import] Roster sheet 'Current_Students' empty or header only.");
    } else {
      const header = rosterData[0].map(h => String(h).toLowerCase().trim());
      const nameCol = header.indexOf("student_name");
      const folderIdCol = header.indexOf("drive_folder_id");
      const emailCol = header.indexOf("student_email");


      let missingCols = [];
      if (nameCol === -1) missingCols.push('"student_name"');
      if (folderIdCol === -1) missingCols.push('"drive_folder_id"');
      if (emailCol === -1) missingCols.push('"student_email"');


      if (missingCols.length > 0) {
        Logger.log(`[ERROR Import] Missing Roster columns: ${missingCols.join(', ')}`);
        return;
      }


      for (let i = 1; i < rosterData.length; i++) {
        const nameRaw = rosterData[i][nameCol];
        const idRaw = rosterData[i][folderIdCol];
        const emailRaw = rosterData[i][emailCol];
        if (nameRaw?.toString().trim() && idRaw?.toString().trim() && emailRaw?.toString().trim()) {
          const nameProcessed = nameRaw.toString().toLowerCase().replace(/_/g, " ").trim();
          if (!rosterMap.has(nameProcessed)) {
            rosterMap.set(nameProcessed, {
                folderId: idRaw.toString().trim(),
                originalName: nameRaw.toString().trim(),
                email: emailRaw.toString().trim()
            });
          } // else { Logger.log(`[WARN Import] Duplicate name: "${nameProcessed}"`); } // Optional log
        } // else { Logger.log(`[WARN Import] Skipping Roster row ${i + 1} missing data.`); } // Optional log
      }
    }
    Logger.log(`[Import] Roster map created with ${rosterMap.size} entries.`);
  } catch (e) {
    Logger.log(`[ERROR Import] Reading Roster: ${e.message}`); return;
  }
  if (!rosterMap || rosterMap.size === 0) { Logger.log("[WARN Import] Roster map empty."); return; }


  // --- List Transcripts in GCS ---
  const textMimeType = ['text/plain'];
  let transcripts;
  try {
    transcripts = listGCSFiles(currentBucketName, "transcripts/", textMimeType);
  } catch (listError) { Logger.log(`[ERROR Import] Listing GCS files: ${listError.message}`); return; }
  Logger.log(`[Import] Found ${transcripts.length} text files in GCS path "transcripts/".`);
  if (transcripts.length === 0) { Logger.log("[Import] No new transcripts found."); return; }


  // --- Process Each Transcript ---
  let importCount = 0;
  let errorCount = 0;
  let skippedExistsCount = 0;
  let skippedNoMatchCount = 0;
  let skippedAsteriskCount = 0;


  // *** Map to store successfully imported files per student for this run ***
  const importedFilesByStudent = new Map(); // Key: studentEmail, Value: Array of transcript filenames


  transcripts.forEach(transcriptItem => {
    // Logger.log(`[DEBUG Import] Processing item: ${JSON.stringify(transcriptItem)}`); // Keep if needed
    try {
      if (!transcriptItem?.name) { errorCount++; return; }
      const fullGcsPath = transcriptItem.name;
      if (!fullGcsPath.startsWith('transcripts/') || fullGcsPath.endsWith('/') || fullGcsPath.substring('transcripts/'.length).includes('/')) return;
      const fileNameRelative = fullGcsPath.substring('transcripts/'.length);
      if (!fileNameRelative) return;


      const match = fileNameRelative.match(/^(.+?)_(\d{4}-\d{2}-\d{2})(?:_[A-Za-z0-9_-]{10})?\.txt$/i);
      if (!match?.[1]) { Logger.log(`[WARN Import] Skipping unrecognized format: "${fileNameRelative}"`); return; }
      let studentNamePartRaw = match[1];
      if (studentNamePartRaw.includes('*')) { skippedAsteriskCount++; return; }
      const warningCharsRegex = /[+\-~`']/g;
      let cleanedNamePart = studentNamePartRaw.replace(warningCharsRegex, ' ');
      const studentNameProcessed = cleanedNamePart.replace(/_/g, ' ').replace(/\s+/g, ' ').toLowerCase().trim();
      if (!studentNameProcessed) { skippedNoMatchCount++; return; }


      if (!rosterMap.has(studentNameProcessed)) { skippedNoMatchCount++; return; }
      const studentData = rosterMap.get(studentNameProcessed);
      const { folderId, originalName: matchedRosterName, email: studentEmail } = studentData;
      if (!studentEmail) { Logger.log(`[WARN Import] Email missing for ${matchedRosterName}.`); return; }


      let transcriptContent = fetchGCSFileContent(currentBucketName, fullGcsPath, accessToken);
      if (transcriptContent === null) { throw new Error(`GCS fetch failed: ${fullGcsPath}`); }


      const targetFolder = DriveApp.getFolderById(folderId);
      const existingFiles = targetFolder.getFilesByName(fileNameRelative);


      if (existingFiles.hasNext()) {
        skippedExistsCount++;
      } else {
        const newFile = targetFolder.createFile(fileNameRelative, transcriptContent, MimeType.PLAIN_TEXT);
        Logger.log(`[Import] SUCCESS: Uploaded "${newFile.getName()}" to Drive for "${matchedRosterName}".`);
        importCount++;


        // --- >>> Store filename for later homework assignment <<< ---
        if (!importedFilesByStudent.has(studentEmail)) {
            importedFilesByStudent.set(studentEmail, []);
        }
        importedFilesByStudent.get(studentEmail).push(fileNameRelative);
        Logger.log(`[Import] Queued "${fileNameRelative}" for homework assignment for ${studentEmail}.`);
      }
    } catch (error) {
      errorCount++;
      const errorFileName = transcriptItem?.name || "unknown file";
      Logger.log(`[ERROR Import] Processing transcript "${errorFileName}". Error: ${error.message}`);
      if (error.message.toLowerCase().includes("not found") && error.message.toLowerCase().includes("folder")) Logger.log(`   Suggestion: Verify Folder ID ${folderId} exists and SA has access.`);
      else if (error.message.toLowerCase().includes("permission") || error.message.toLowerCase().includes("forbidden")) Logger.log(`   Suggestion: Check SA Editor permission on Drive Folder ID ${folderId}.`);
      // Add other specific suggestions if needed
      if (error.stack) Logger.log(`   Stack: ${error.stack}`);
    }
  }); // End loop through GCS transcripts


  // --- Assign Homework AFTER looping through all files ---
  Logger.log(`[Import] Finished GCS scan. Attempting to assign homework for ${importedFilesByStudent.size} students...`);
  let assignmentErrors = 0;
  for (const [studentEmail, transcriptFileNames] of importedFilesByStudent.entries()) {
      if (transcriptFileNames && transcriptFileNames.length > 0) {
          try {
              Logger.log(`[Import] Assigning homework to ${studentEmail} using ${transcriptFileNames.length} transcript(s): ${transcriptFileNames.join(', ')}`);
              assignHomework(studentEmail, transcriptFileNames); // Pass the array
              Logger.log(`[Import] Homework assignment successful for ${studentEmail}.`);
          } catch (assignError) {
              Logger.log(`[ERROR Import] Failed assignHomework call for ${studentEmail} files [${transcriptFileNames.join(', ')}]: ${assignError.message}`);
              assignmentErrors++;
          }
      }
  }


  Logger.log(`[Import] Finished. Imported: ${importCount}, Skipped (Exists): ${skippedExistsCount}, Skipped (No Match): ${skippedNoMatchCount}, Skipped (*): ${skippedAsteriskCount}, Processing Errors: ${errorCount}, Assignment Errors: ${assignmentErrors}.`);
}




/**
 * =====================================================================
 *  EDU SCRIBE  –  HOMEWORK PROMPT UTILITIES  (May 2025)
 * =====================================================================
 */


/* ── column indices in ‘Current_Students’ (0-based) ────────────────── */
const COL_ROSTER_STUDENT_NAME       = 0;   // A: Student_Name
const COL_ROSTER_STUDENT_ID         = 1;   // B: Student_ID
const COL_ROSTER_LIFE_AND_LIFESTYLE = 16;  // Q: Life_And_Liestyle
/* Drive_Folder_ID index is read from Script Property ROSTER_DRIVE_FOLDER_ID */
/* Student_Email column is found dynamically by header name           */


/* ── global IDs pulled once from Script Properties ─────────────────── */
const SPREAD_ID       = PropertiesService.getScriptProperties().getProperty('STUDENT_ROSTER_ID');
const HW_SHEET_NAME   = PropertiesService.getScriptProperties().getProperty('HOMEWORK_SHEET_NAME');
const DRIVE_COL_PROP  = PropertiesService.getScriptProperties().getProperty('ROSTER_DRIVE_FOLDER_ID');


/* ==================================================================== */
/*  rosterFindByEmail – returns student object incl. LifeStyle          */
/* ==================================================================== */
/**
 * Looks up a student in the **Current_Students** tab by their email address
 * and returns a rich object *including* the Life & Lifestyle text (column Q)
 * and Drive folder ID (column index now stored as a Script Property).
 */
function rosterFindByEmail(email) {
  if (!email) { Logger.log('[rosterFindByEmail] Called with null or empty email.'); return null; }
  if (!SPREAD_ID) { Logger.log('[rosterFindByEmail] SPREAD_ID global variable is not set.'); return null; }
  let ss; try { ss = SpreadsheetApp.openById(SPREAD_ID); }
  catch (e) { Logger.log(`[rosterFindByEmail] Failed to open Spreadsheet ID: ${SPREAD_ID}. Error: ${e.message}`); return null; }
  const sheet = ss.getSheetByName('Current_Students');
  if (!sheet) { Logger.log(`[rosterFindByEmail] Sheet 'Current_Students' not found in Spreadsheet ID: ${SPREAD_ID}.`); return null; }
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) { Logger.log('[rosterFindByEmail] Sheet "Current_Students" is empty or has only a header.'); return null; }


  const headerLower = values[0].map(h => String(h).toLowerCase().trim());
  const emailColIdx = headerLower.indexOf('student_email');
  if (emailColIdx === -1) { Logger.log('[rosterFindByEmail] Column "Student_Email" not found.'); return null; }


  let driveColIdx = -1;
  if (DRIVE_COL_PROP !== null && DRIVE_COL_PROP.trim() !== '' && !isNaN(DRIVE_COL_PROP)) { driveColIdx = Number(DRIVE_COL_PROP); }
  else { Logger.log('[rosterFindByEmail] Warning: Script Property "ROSTER_DRIVE_FOLDER_ID" missing/invalid.'); return null; }


  const target = email.toLowerCase().trim();
  for (let r = 1; r < values.length; r++) {
    const row = values[r]; const rowEmail = row[emailColIdx];
    if (rowEmail && typeof rowEmail === 'string' && rowEmail.trim().toLowerCase() === target) {
       const studentId = (COL_ROSTER_STUDENT_ID >= 0 && COL_ROSTER_STUDENT_ID < row.length) ? row[COL_ROSTER_STUDENT_ID] : null;
       const studentName = (COL_ROSTER_STUDENT_NAME >= 0 && COL_ROSTER_STUDENT_NAME < row.length) ? row[COL_ROSTER_STUDENT_NAME] : null;
       const driveFolderId = (driveColIdx >= 0 && driveColIdx < row.length) ? row[driveColIdx] : null;
       const lifeStyle = (COL_ROSTER_LIFE_AND_LIFESTYLE >= 0 && COL_ROSTER_LIFE_AND_LIFESTYLE < row.length) ? row[COL_ROSTER_LIFE_AND_LIFESTYLE] : '';
       if (driveFolderId === null) Logger.log(`[rosterFindByEmail] Warning: Missing Drive_Folder_ID for ${email} row ${r+1}`);


       return {
        Student_ID       : studentId?.toString() || '',
        Name             : studentName?.toString() || '',
        Email            : rowEmail.trim(),
        Drive_Folder_ID  : driveFolderId?.toString() || '',
        LifeStyle        : lifeStyle?.toString() || '',
        RosterRowIndex   : r + 1
      };
    }
  }
  Logger.log(`[rosterFindByEmail] Email "${email}" not found.`); return null;
}




/* ==================================================================== */
/*  saveHomeworkPrompt – stores prompt file + ledger row                */
/* ==================================================================== */
/**
 * V3: (Based on original "V2" structure from active project) Accepts pre-constructed prompt filename,
 *     writes "Active" to Token_Status, and initializes "Turns_Used" to 0.
 * Creates a prompt file in the student’s Drive folder using the provided filename,
 * appends a row to the Homework_Push ledger (including "Active" status and 0 Turns_Used),
 * and returns the token.
 */
function saveHomeworkPrompt(studentEmail, hwId, basePromptText, promptFileName) {
  const student = rosterFindByEmail(studentEmail);
  if (!student) {
    Logger.log(`[saveHomeworkPrompt] ERROR: Student email ${studentEmail} not found.`); // Standardized error log
    throw new Error(`Student email ${studentEmail} not found.`);
  }
  if (!student.Drive_Folder_ID) {
    Logger.log(`[saveHomeworkPrompt] ERROR: Drive_Folder_ID missing for student ${studentEmail}.`); // Standardized error log
    throw new Error(`Drive_Folder_ID missing for student ${studentEmail}.`);
  }
  if (!promptFileName?.trim()) {
      const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
      promptFileName = `HW_${hwId || `Unknown-${today}`}_prompt.txt`;
      Logger.log(`[saveHomeworkPrompt] WARNING: Invalid promptFileName. Using fallback: ${promptFileName}`);
  }


  const fullPrompt = `${basePromptText.trim()}\n\n---\nStudent background / Life & Lifestyle\n${(student.LifeStyle || 'No profile data on record.').trim()}`;
  let promptFile;
  try {
    const targetFolder = DriveApp.getFolderById(student.Drive_Folder_ID);
    promptFile = targetFolder.createFile(promptFileName, fullPrompt, MimeType.PLAIN_TEXT);
    Logger.log(`[saveHomeworkPrompt] Created prompt file "${promptFileName}" (ID: ${promptFile.getId()}) in folder ID ${student.Drive_Folder_ID} for student ${studentEmail}.`); // Standardized log
  } catch (e) {
    Logger.log(`[saveHomeworkPrompt] ERROR creating prompt file "${promptFileName}" for student ${studentEmail}: ${e.message}`); // Standardized log
    throw new Error(`Failed to create prompt file: ${e.message}`);
  }


  const token = Utilities.getUuid();
  try {
    const currentSpreadId = PropertiesService.getScriptProperties().getProperty('STUDENT_ROSTER_ID');
    const currentHwSheetName = PropertiesService.getScriptProperties().getProperty('HOMEWORK_SHEET_NAME'); // This is "Homework_Push"


    if (!currentSpreadId) throw new Error("STUDENT_ROSTER_ID script property missing."); // Standardized error
    if (!currentHwSheetName) throw new Error("HOMEWORK_SHEET_NAME script property missing for Homework_Push."); // Standardized error
   
    const ledgerSheet = SpreadsheetApp.openById(currentSpreadId).getSheetByName(currentHwSheetName);
    if (!ledgerSheet) throw new Error(`Ledger sheet "${currentHwSheetName}" not found.`);


    // Ensure your appendRow array matches the exact order and number of columns in your "Homework_Push" sheet
    // Headers: Student_ID, Student_Name, Student_Email, HW_ID, PromptFileID, Token, AssignedAt, CompletedAt, Token_Status, Turns_Used
    const newRowData = [
      student.Student_ID || '',       // Column 1 (A)
      student.Name || '',             // Column 2 (B)
      student.Email || '',            // Column 3 (C)
      hwId || '',                     // Column 4 (D)
      promptFile.getId(),             // Column 5 (E)
      token,                          // Column 6 (F)
      new Date().toISOString(),       // Column 7 (G) (AssignedAt)
      '',                             // Column 8 (H) (CompletedAt - initially blank)
      'Active',                       // Column 9 (I) (Token_Status - SET TO ACTIVE)
      0                               // Column 10 (J) (Turns_Used - INITIALIZE TO 0)
    ];
   
    ledgerSheet.appendRow(newRowData);
    SpreadsheetApp.flush(); // Ensure changes are written
    Logger.log(`[saveHomeworkPrompt] Successfully appended to "${currentHwSheetName}". HW_ID: ${hwId}, Token: ${token}, Status: Active, Turns_Used: 0, for ${student.Name}.`);


  } catch (e) {
    Logger.log(`[saveHomeworkPrompt] ERROR appending ledger row for HW_ID ${hwId}, student ${studentEmail}: ${e.message}`); // Standardized log
    try {
      if (promptFile) {
        promptFile.setTrashed(true);
        Logger.log(`[saveHomeworkPrompt] Cleaned up (trashed) prompt file ${promptFile.getId()} due to ledger error.`); // Standardized log
      }
    } catch (trashErr) {
      Logger.log(`[saveHomeworkPrompt] WARNING: Failed to trash prompt file ${promptFile.getId()} after ledger error: ${trashErr.message}`);
    }
    throw new Error(`Failed to append ledger row: ${e.message}`);
  }
 
  Logger.log(`[saveHomeworkPrompt] HW assignment ${hwId} processed for ${student.Name}. Prompt File: "${promptFileName}". Token: ${token}.`);
  return token;
}




/* ==================================================================== */
/*  assignHomework – triggers prompt generation and email               */
/* ==================================================================== */
/**
 * V3: Accepts an array of transcript filenames, uses the first for prompt naming,
 *     and lists all in the prompt. Email body text modified.
 */
function assignHomework(studentEmail, transcriptFileNames) {
  if (!studentEmail || !Array.isArray(transcriptFileNames) || transcriptFileNames.length === 0) { throw new Error("Missing parameters for assignHomework."); }
  const student = rosterFindByEmail(studentEmail);
  if (!student) { Logger.log(`[assignHomework] Student not found: ${studentEmail}.`); throw new Error(`Student lookup failed for ${studentEmail}`); }


  const firstTranscriptName = transcriptFileNames[0];
  const transcriptMatch = firstTranscriptName.match(/_(\d{4}-\d{2}-\d{2})_([A-Za-z0-9_-]{10})\.txt$/i);
  let promptFileName = null;
  if (transcriptMatch?.[1] && transcriptMatch?.[2]) { promptFileName = `prompt_${transcriptMatch[1]}_${transcriptMatch[2]}.txt`; Logger.log(`[assignHomework] Prompt filename: ${promptFileName}`); }
  else { Logger.log(`[assignHomework] WARNING: Cannot parse "${firstTranscriptName}". Using default naming.`); const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd'); const fallbackHwId = `L${student.RosterRowIndex || 'X'}-${today}`; promptFileName = `HW_${fallbackHwId}_prompt.txt`; }


  const todayForHwId = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
  const hwId = `L${student.RosterRowIndex || 'X'}-${todayForHwId}`;
  const transcriptList = transcriptFileNames.join(', ');
  const basePrompt = `You are a GPT called Homework Coach.\nTranscript files for this session: ${transcriptList}.\nThe exact grammar topic being studied might be mentioned directly in the conversation. Search for this to identify what it is, but if more than one grammar topic or grammar theme is mentioned, then surmize the grammar topic being taught, judging from repeated sentence structures being practiced and especially how the teacher introduces the topic and corrects the student. Carefully analyze the student’s communications and reactions to the teacher’s instructions to find instances in which the grammar being taught was not well comprehended by the student and based on this analysis, choosing only one grammatical theme to pursue and formulate all ten questions and dialogue around this one theme for this homework session. In your greeting, explicitly state the grammar topic being covered in the homework, followed by a very brief situation based on the student's life and lifestyle data or a relevant comment from the student transcript to provide context showing how the grammar structure can be productively applied. Ask up to 10 personalized questions based on the conclusions you reached in your analysis of the transcript and incorporate the student background provided below. Show the problem/question number (i) for each question asked. After student answers, provide an instant, short and empathetic evaluation of the student's answer, explaining clearly but briefly why the student is right or what needs work, then in the same output, move on to the next student prompt, with the next number (i+1) displayed for reference. Focus on reinforcing weak points with this grammar topic identified in the transcript(s), being sure to add professional and personal details from the transcript and the student information below to keep the output relevant to the student’s profession, personal characteristics and world view. General behavior when interacting with the student: *Do not prompt the student to SPEAK or LISTEN to you. You will be interacting by text chat only so speaking and listening is not possible. *Be friendly and empathetic but brief in your responses. *When you’ve finished coaching the student through the 10 homework problems, it is essential to bring the conversation to a warm and polite close by instructing the student to click the button that says ‘Done ✅ Submit Homework’.`;
  const token = saveHomeworkPrompt(studentEmail, hwId, basePrompt, promptFileName);


  const portalBaseUrl = PropertiesService.getScriptProperties().getProperty('HOMEWORK_PORTAL_BASEURL');
  if (!portalBaseUrl) { Logger.log("[assignHomework] WARNING: HOMEWORK_PORTAL_BASEURL property not set. Skipping email."); return; }
 
  // --- ENSURING THE CORRECT PORTAL LINK STRUCTURE AS PER LATEST INSTRUCTION ---
  const portalLink = `${portalBaseUrl}/homework-coach/?token=${token}`; // Removed /index.php
  // --- END PORTAL LINK CORRECTION ---


  const studentFirstName = student.Name ? student.Name.split(' ')[0] : "Student";
  // const emailTranscriptRef = transcriptFileNames.length > 1 ? `${transcriptFileNames.length} transcripts from your last session` : `transcript ${firstTranscriptName}`; // This variable is no longer used in the email body


  try {
    MailApp.sendEmail( studentEmail, `Your new homework is ready (ID: ${hwId})`, `Hi ${studentFirstName},\n\nYour practice set based on your last class is ready. Click the link below to open it in your portal:\n\n${portalLink}\n\nRemember: complete all ten turns with the Homework GPT, then hit Done ✅ to turn in your homework.\n\nGood luck!\n` );
    Logger.log(`[assignHomework] HW ${hwId} email sent to ${studentEmail}`);
  } catch (e) { Logger.log(`[assignHomework] ERROR sending email to ${studentEmail} for HW ${hwId}: ${e.message}`); }
}




// ------------------------------
// GCS File Listing Utility
// ------------------------------
/** Lists files in GCS bucket. V2 uses startsWith for mime type. */
function listGCSFiles(bucketName, prefix, mimeTypes) {
  let accessToken; try { accessToken = getServiceAccountToken(); }
  catch (e) { throw new Error(`Auth failed in listGCSFiles: ${e.message}`); }
  if (!bucketName) throw new Error("Bucket name required.");


  const effectivePrefix = (prefix && prefix !== "/" && !prefix.endsWith('/')) ? `${prefix}/` : (prefix === "/" ? "" : prefix);
  let url = `https://storage.googleapis.com/storage/v1/b/${bucketName}/o?prefix=${encodeURIComponent(effectivePrefix)}&fields=items(name,contentType,size),nextPageToken`;
  let allFiles = []; let pageToken = null;
  // Logger.log(`[ListGCS DEBUG] Listing gs://${bucketName}/${effectivePrefix}`); // Optional verbose log


  try {
    do {
      let currentUrl = url; if (pageToken) currentUrl += `&pageToken=${pageToken}`;
      let response = UrlFetchApp.fetch(currentUrl, { method: 'GET', headers: { 'Authorization': `Bearer ${accessToken}` }, muteHttpExceptions: true });
      let responseCode = response.getResponseCode(); let responseText = response.getContentText();
      if (responseCode !== 200) { Logger.log(`[ERROR ListGCS] HTTP ${responseCode} prefix "${effectivePrefix}". ${responseText}`); let googleError = responseText; try { googleError = JSON.parse(responseText).error.message || responseText; } catch(ignore) {} throw new Error(`GCS list files failed (${responseCode}): ${googleError}`); }
      let result = JSON.parse(responseText);
      if (result.items) {
        const filteredItems = result.items.filter(item => {
          if (effectivePrefix && item.name === effectivePrefix && item.contentType === 'application/x-www-form-urlencoded;charset=UTF-8') return false; // Skip folder object
          if (!item.name) return false;
          if (mimeTypes?.length > 0) { const matchesMime = mimeTypes.some(type => item.contentType && item.contentType.toLowerCase().startsWith(type.toLowerCase())); if (!matchesMime) return false; }
          return true;
        });
        allFiles = allFiles.concat(filteredItems);
      }
      pageToken = result.nextPageToken; 
    } while (pageToken);
  } catch (e) { Logger.log(`[ERROR ListGCS] Exc listing "${effectivePrefix}": ${e.message}`); throw new Error(`Exception listing GCS files: ${e.message}`); }
  // Logger.log(`[ListGCS DEBUG] Found ${allFiles.length} files.`); // Optional verbose log
  return allFiles;
}




// ------------------------------
// TESTING FUNCTIONS
// ------------------------------


function testAuthToken() { 
  try {
    const token = getServiceAccountToken();
    Logger.log(`SUCCESS: Token retrieved: ${token.substring(0, 30)}...`);
  } catch (e) {
    Logger.log(`ERROR: ${e.message}`);
  }
}
function testStorageAccess() { 
  const props = PropertiesService.getScriptProperties();
  const bucketName = props.getProperty("BUCKET_NAME");
  if (!bucketName) { Logger.log("ERROR: BUCKET_NAME not set in Script Properties."); return; }
  try {
    const token = getServiceAccountToken();
    const url = `https://storage.googleapis.com/storage/v1/b/${bucketName}?fields=name`;
    const response = UrlFetchApp.fetch(url, {
      headers: { Authorization: `Bearer ${token}` },
      muteHttpExceptions: true,
    });
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    if (responseCode === 200) {
      Logger.log(`SUCCESS: Able to access bucket details for "${bucketName}". Response: ${responseText}`);
    } else {
      Logger.log(`ERROR: Failed to access bucket "${bucketName}". Code: ${responseCode}. Response: ${responseText}`);
    }
  } catch (e) {
    Logger.log(`ERROR: Exception during storage access test: ${e.message}`);
  }
}
function testListGCSIncomingVideos() { 
  const props = PropertiesService.getScriptProperties();
  const bucketName = props.getProperty("BUCKET_NAME");
   if (!bucketName) { Logger.log("ERROR: BUCKET_NAME not set."); return; }
  try {
    const files = listGCSFiles(bucketName, "incoming/", ["video/mp4"]);
    if (files && files.length > 0) {
      Logger.log(`Found ${files.length} MP4 files in gs://${bucketName}/incoming/:`);
      files.forEach(file => Logger.log(` - ${file.name} (Size: ${file.size}, Type: ${file.contentType})`));
    } else if (files) {
      Logger.log(`No MP4 files found in gs://${bucketName}/incoming/.`);
    } else {
      Logger.log(`listGCSFiles returned null or undefined for gs://${bucketName}/incoming/.`);
    }
  } catch (e) {
    Logger.log(`Error listing incoming videos: ${e.message}\nStack: ${e.stack}`);
  }
}
function testListGCSTranscripts() { 
  const props = PropertiesService.getScriptProperties();
  const bucketName = props.getProperty("BUCKET_NAME");
  if (!bucketName) { Logger.log("ERROR: BUCKET_NAME not set."); return; }
  try {
    const files = listGCSFiles(bucketName, "transcripts/", ["text/plain"]);
     if (files && files.length > 0) {
      Logger.log(`Found ${files.length} TXT files in gs://${bucketName}/transcripts/:`);
      files.forEach(file => Logger.log(` - ${file.name} (Size: ${file.size}, Type: ${file.contentType})`));
    } else if (files) {
      Logger.log(`No TXT files found in gs://${bucketName}/transcripts/.`);
    } else {
      Logger.log(`listGCSFiles returned null or undefined for gs://${bucketName}/transcripts/.`);
    }
  } catch (e) {
    Logger.log(`Error listing transcripts: ${e.message}\nStack: ${e.stack}`);
  }
}
function testImportTranscripts() { 
  Logger.log("Starting testImportTranscripts manually...");
  importTranscriptsToDrive();
  Logger.log("Finished testImportTranscripts.");
}
function testFileAccess() { 
  const props = PropertiesService.getScriptProperties();
  const folderId = props.getProperty("FOLDER_ID");
  if (!folderId) { Logger.log("ERROR: FOLDER_ID not set."); return; }
  try {
    const folder = DriveApp.getFolderById(folderId);
    Logger.log(`Successfully accessed folder: "${folder.getName()}".`);
    const files = folder.getFiles();
    let count = 0;
    while (files.hasNext() && count < 5) {
      const file = files.next();
      Logger.log(` - File: ${file.getName()} (MIME: ${file.getMimeType()})`);
      count++;
    }
    if (count === 0) Logger.log("No files found in the folder or folder is empty.");
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
    Logger.log(`Successfully accessed roster: "${ss.getName()}", sheet: "${sheet.getName()}".`);
    const firstRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    Logger.log(`Header row: ${firstRow.join(", ")}`);
    const numStudents = sheet.getLastRow() -1;
    Logger.log(`Number of student entries (excluding header): ${numStudents}`);
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
function testHandleCompletedJobManually() { 
  // --- CONFIGURATION FOR MANUAL TEST ---
  const testFileName = "REPLACE_WITH_REAL_FILENAME_FROM_GCS_INCOMING.mp4"; // e.g., "Student_Name_2023-10-26_abc123def4.mp4"
  const testJobId = "REPLACE_WITH_REAL_SPEECHMATICS_JOB_ID"; // e.g., "xxxxxxxx"
  // --- END CONFIGURATION ---

  if (testFileName.startsWith("REPLACE_") || testJobId.startsWith("REPLACE_")) {
    Logger.log("ERROR: Please configure testFileName and testJobId in testHandleCompletedJobManually first.");
    return;
  }

  const props = PropertiesService.getScriptProperties();
  const bucketName = props.getProperty("BUCKET_NAME");
  const apiKey = props.getProperty("SPEECHMATICS_API_KEY");

  if (!bucketName || !apiKey) {
    Logger.log("ERROR: BUCKET_NAME or SPEECHMATICS_API_KEY missing from Script Properties.");
    return;
  }

  Logger.log(`[Manual Test] Testing handleCompletedJob for File: "${testFileName}", Job ID: "${testJobId}"`);
  let accessToken;
  try {
    accessToken = getServiceAccountToken();
    Logger.log("[Manual Test] Service Account Token obtained.");
  } catch (e) {
    Logger.log(`[Manual Test] ERROR: Failed to get Service Account Token: ${e.message}`);
    return;
  }

  try {
    handleCompletedJob(testFileName, testJobId, bucketName, accessToken, apiKey);
    Logger.log(`[Manual Test] handleCompletedJob executed for "${testFileName}". Check GCS "transcripts/" folder.`);

    // Optional: Test moving the original file in GCS
    const sourceGCSPath = `incoming/${testFileName}`;
    const destinationGCSPath = `processed/${testFileName}`;
    Logger.log(`[Manual Test] Attempting to move "${sourceGCSPath}" to "${destinationGCSPath}" in GCS bucket "${bucketName}".`);
    moveFileInGCS(bucketName, sourceGCSPath, destinationGCSPath, accessToken);
    Logger.log(`[Manual Test] GCS move operation attempted. Check GCS "incoming/" and "processed/" folders.`);

  } catch (e) {
    Logger.log(`[Manual Test] ERROR during handleCompletedJob or GCS move: ${e.message}\nStack: ${e.stack || 'N/A'}`);
  }
  Logger.log("[Manual Test] Finished.");
}


// --- Homework Utility Test Functions ---


/** Temporary wrapper to test saveHomeworkPrompt directly */
function testSaveHomeworkPrompt() {
  // 📧 !!! REPLACE with a real email from 'Current_Students' !!!
  const testEmail = 'mateo.foster1@gmail.com';
  // --- End Replace ---
  const testHwId = 'L0-Test-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMddHHmmss");
  const testBasePrompt = 'This is a test base prompt.';
  const dummyDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  const dummySuffix = Utilities.getUuid().substring(0, 10).replace(/-/g, '_');
  const testPromptFileName = `prompt_${dummyDate}_${dummySuffix}.txt`;


  if (testEmail.startsWith('replace.me')) { Logger.log("ERROR: Please replace 'replace.me@example.com' first."); return; }


  Logger.log(`[DEBUG] === Testing saveHomeworkPrompt for ${testEmail} ===`);
  Logger.log(`[DEBUG]   Using HW_ID: ${testHwId}`); Logger.log(`[DEBUG]   Using Prompt Filename: ${testPromptFileName}`);
  try {
    const token = saveHomeworkPrompt(testEmail, testHwId, testBasePrompt, testPromptFileName);
    Logger.log(`[SUCCESS] saveHomeworkPrompt executed. Token: ${token}`);
    Logger.log(`   Check Drive folder for ${testEmail} for file "${testPromptFileName}"`);
    Logger.log(`   Check "Homework_Push" sheet for row with HW_ID ${testHwId} and Turns_Used = 0.`); // Updated log
  } catch (e) { Logger.log(`[ERROR] Test failed: ${e.message}\n   Stack: ${e.stack || 'N/A'}`); }
  Logger.log("[DEBUG] === Finished testing saveHomeworkPrompt ===");
}


/** Temporary wrapper to test assignHomework directly */
function testAssignHomework() {
  // 📧 !!! REPLACE with a real email from 'Current_Students' !!!
  const testEmail = 'mateo.foster1@gmail.com';
  // --- End Replace ---
  const dummyDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  const dummySuffix1 = Utilities.getUuid().substring(0, 10).replace(/-/g, '_');
  const dummySuffix2 = Utilities.getUuid().substring(0, 10).replace(/-/g, '_');
  const testTranscriptFileNames = [ `FakeyMcFakerson_${dummyDate}_${dummySuffix1}.txt`, `FakeyMcFakerson_${dummyDate}_${dummySuffix2}.txt` ];


  if (testEmail.startsWith('replace.me')) { Logger.log("ERROR: Please replace 'replace.me@example.com' first."); return; }


   Logger.log(`[DEBUG] === Testing assignHomework for ${testEmail} ===`);
   Logger.log(`[DEBUG]   Using Transcript Filenames: [${testTranscriptFileNames.join(', ')}]`);
   try {
       const portalUrl = PropertiesService.getScriptProperties().getProperty('HOMEWORK_PORTAL_BASEURL');
       if (!portalUrl) Logger.log("[ERROR] Property 'HOMEWORK_PORTAL_BASEURL' not set.");
       assignHomework(testEmail, testTranscriptFileNames);
       Logger.log(`[SUCCESS] assignHomework executed for ${testEmail}.`);
       Logger.log(`   Check Drive folder for student ${testEmail} for prompt file.`);
       Logger.log(`   Check "Homework_Push" sheet for ledger row (should include Turns_Used = 0).`); // Updated log
       Logger.log(`   Check inbox for ${testEmail} for notification email.`);
   } catch (e) { Logger.log(`[ERROR] Test failed: ${e.message}\n   Stack: ${e.stack || 'N/A'}`); }
   Logger.log("[DEBUG] === Finished testing assignHomework ===");
}


/**
 * ========== END COMPLETE SCRIPT ==========
 */
