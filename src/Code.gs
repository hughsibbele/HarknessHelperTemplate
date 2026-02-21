/**
 * Harkness Discussion Helper — Main Entry Point
 *
 * Pipeline:
 * 1. Monitor Drive folder for uploaded audio recordings
 * 2. Transcribe with ElevenLabs (synchronous, with speaker diarization)
 * 3. Gemini auto-suggests speaker mapping from introductions
 * 4. Teacher reviews speaker map (individual mode) or auto-advances (group mode)
 * 5. Teacher enters grade(s), clicks "Generate Feedback" → Gemini generates
 * 6. Teacher approves → clicks "Send" → email and/or Canvas distribution
 *
 * Status machine: uploaded → transcribing → mapping → review → approved → sent
 *                                                                        ↘ error
 */

// ============================================================================
// CONSTANTS
// ============================================================================

/** Auto-stop processing after this many minutes */
const PROCESSING_TIMEOUT_MINUTES = 60;

// ============================================================================
// MENU & UI
// ============================================================================

/**
 * Create custom menu when spreadsheet opens.
 * Shows minimal menu before setup, full menu after.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  if (!isSetupComplete()) {
    ui.createMenu('Harkness Helper')
      .addItem('Setup Wizard (start here)', 'showSetupWizard')
      .addToUi();
    return;
  }

  ui.createMenu('Harkness Helper')
    .addItem('Start Processing', 'menuStartProcessing')
    .addItem('Stop Processing', 'menuStopProcessing')
    .addSeparator()
    .addItem('Generate Feedback', 'menuGenerateFeedback')
    .addItem('Send Approved Feedback', 'menuSendApprovedFeedback')
    .addSeparator()
    .addItem('Configure Canvas Course', 'showCanvasConfigDialog')
    .addSeparator()
    .addItem('Re-run Setup Wizard', 'showSetupWizard')
    .addToUi();
}

// ============================================================================
// SETUP WIZARD
// ============================================================================

/**
 * Show the Setup Wizard dialog.
 * Collects ElevenLabs and Gemini API keys, then auto-configures everything.
 */
function showSetupWizard() {
  const props = PropertiesService.getScriptProperties().getProperties();

  // Escape HTML special characters to prevent XSS via stored property values
  const escapeHtml = (str) => String(str || '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');

  const existingElevenlabs = escapeHtml(props['ELEVENLABS_API_KEY']);
  const existingGemini = escapeHtml(props['GEMINI_API_KEY']);

  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; color: #333; }
      h2 { color: #1a73e8; margin-top: 0; }
      label { display: block; margin-top: 16px; font-weight: bold; font-size: 14px; }
      input { width: 100%; padding: 8px; margin-top: 4px; border: 1px solid #ccc;
              border-radius: 4px; font-size: 14px; box-sizing: border-box; }
      .help { font-size: 12px; color: #666; margin-top: 2px; }
      button { margin-top: 24px; padding: 10px 24px; background: #1a73e8; color: white;
               border: none; border-radius: 4px; font-size: 14px; cursor: pointer; width: 100%; }
      button:hover { background: #1557b0; }
      button:disabled { background: #ccc; cursor: default; }
      .error { color: #d93025; margin-top: 8px; font-size: 13px; }
      .success { color: #188038; margin-top: 8px; font-size: 13px; }
    </style>

    <h2>Harkness Helper Setup</h2>
    <p>Enter your API keys below. Everything else is configured automatically.</p>

    <label for="elevenlabs">ElevenLabs API Key</label>
    <input type="password" id="elevenlabs" value="${existingElevenlabs}"
           placeholder="Enter your ElevenLabs API key">
    <div class="help">Get one at <b>elevenlabs.io</b> → Profile → API Keys</div>

    <label for="gemini">Gemini API Key</label>
    <input type="password" id="gemini" value="${existingGemini}"
           placeholder="Enter your Gemini API key">
    <div class="help">Get one at <b>aistudio.google.com</b> → API Keys</div>

    <div id="status"></div>

    <button id="btn" onclick="runSetup()">Run Setup</button>

    <script>
      function runSetup() {
        var elevenlabs = document.getElementById('elevenlabs').value.trim();
        var gemini = document.getElementById('gemini').value.trim();
        var status = document.getElementById('status');
        var btn = document.getElementById('btn');

        if (!elevenlabs || !gemini) {
          status.className = 'error';
          status.textContent = 'Both API keys are required.';
          return;
        }

        btn.disabled = true;
        btn.textContent = 'Setting up...';
        status.className = '';
        status.textContent = 'Creating folders, initializing sheets...';

        google.script.run
          .withSuccessHandler(function(result) {
            status.className = 'success';
            status.innerHTML = '<b>Setup complete!</b><br><br>' +
              'Next steps:<br>' +
              '1. Upload audio files to the <b>Harkness Helper / Upload</b> folder in your Google Drive<br>' +
              '2. Click <b>Harkness Helper → Start Processing</b><br>' +
              '3. Customize the <b>Settings</b> sheet (mode, distribution, etc.)<br><br>' +
              'Refresh this page to see the full menu.';
            btn.textContent = 'Done!';
          })
          .withFailureHandler(function(err) {
            status.className = 'error';
            status.textContent = 'Error: ' + err.message;
            btn.disabled = false;
            btn.textContent = 'Run Setup';
          })
          .runSetupWizard({ elevenlabs: elevenlabs, gemini: gemini });
      }
    </script>
  `).setWidth(480).setHeight(480);

  SpreadsheetApp.getUi().showModalDialog(html, 'Setup Wizard');
}

/**
 * Server-side handler for the Setup Wizard.
 * Auto-captures spreadsheet ID, stores API keys, creates Drive folders,
 * and initializes all sheets/settings/prompts.
 *
 * @param {Object} config - {elevenlabs, gemini}
 * @returns {Object} {success: true}
 */
function runSetupWizard(config) {
  // 1. Auto-capture spreadsheet ID
  const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
  setProperty('SPREADSHEET_ID', ssId);

  // 2. Store API keys
  setProperty('ELEVENLABS_API_KEY', config.elevenlabs);
  setProperty('GEMINI_API_KEY', config.gemini);

  // 3. Create Drive folders (stores AUDIO_FOLDER_ID and PROCESSING_FOLDER_ID)
  setupDriveFolders();

  // 4. Initialize spreadsheet structure
  initializeAllSheets();
  formatSheets();
  initializeSettings();
  initializeDefaultPrompts();

  Logger.log('Setup Wizard completed successfully');
  return { success: true };
}

// ============================================================================
// AUTO-OFF TRIGGER SYSTEM
// ============================================================================

/**
 * Start processing: install trigger and run immediately.
 * Removes any existing processing triggers first to prevent duplicates.
 */
function menuStartProcessing() {
  const ui = SpreadsheetApp.getUi();

  try {
    // Remove any existing processing triggers
    removeProcessingTriggers();

    // Record start time
    setProperty('PROCESSING_START_TIME', new Date().toISOString());

    // Install 10-minute trigger
    ScriptApp.newTrigger('mainProcessingLoop')
      .timeBased()
      .everyMinutes(CONFIG.TRIGGERS.MAIN_LOOP)
      .create();

    // Run immediately
    mainProcessingLoop();

    ui.alert('Processing Started',
      'Processing is now running and will check for new files every 10 minutes.\n\n' +
      'It will automatically stop after ' + PROCESSING_TIMEOUT_MINUTES + ' minutes.\n\n' +
      'Upload audio files to your Harkness Helper / Upload folder in Google Drive.',
      ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', e.message, ui.ButtonSet.OK);
  }
}

/**
 * Stop processing: remove trigger and clear start time.
 */
function menuStopProcessing() {
  const ui = SpreadsheetApp.getUi();
  removeProcessingTriggers();

  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('PROCESSING_START_TIME');

  ui.alert('Processing Stopped',
    'Automatic processing has been stopped.',
    ui.ButtonSet.OK);
}

/**
 * Remove only mainProcessingLoop triggers (not all project triggers).
 */
function removeProcessingTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'mainProcessingLoop') {
      ScriptApp.deleteTrigger(trigger);
      removed++;
    }
  }
  Logger.log(`Removed ${removed} processing trigger(s)`);
}

// ============================================================================
// MAIN PROCESSING LOOP (runs on 10-minute trigger)
// ============================================================================

/**
 * Main trigger function — runs periodically to process work.
 * Auto-stops after PROCESSING_TIMEOUT_MINUTES.
 */
function mainProcessingLoop() {
  // Check elapsed time — auto-stop if past timeout
  const props = PropertiesService.getScriptProperties();
  const startTime = props.getProperty('PROCESSING_START_TIME');
  if (!startTime) {
    // Orphaned trigger — no start time recorded. Remove it.
    Logger.log('No PROCESSING_START_TIME found. Removing orphaned trigger.');
    removeProcessingTriggers();
    return;
  }
  const elapsed = (Date.now() - new Date(startTime).getTime()) / 60000;
  if (elapsed >= PROCESSING_TIMEOUT_MINUTES) {
    Logger.log(`Processing timeout reached (${Math.round(elapsed)} minutes). Auto-stopping.`);
    removeProcessingTriggers();
    props.deleteProperty('PROCESSING_START_TIME');
    return;
  }

  Logger.log('=== Starting main processing loop ===');

  try {
    // Step 1: Check for new audio files and transcribe
    checkAndProcessNewFiles();

    // Step 2: Detect stuck transcriptions
    checkStuckTranscriptions();

    // Step 3: In group mode, auto-advance confirmed discussions from mapping → review
    advanceMappingStatus();

    Logger.log('=== Main processing loop complete ===');
  } catch (e) {
    Logger.log(`Error in main processing loop: ${e.message}`);
    Logger.log(e.stack);
  }
}

// ============================================================================
// FILE DETECTION & TRANSCRIPTION
// ============================================================================

/**
 * Check for new audio files, transcribe them, and run speaker mapping
 */
function checkAndProcessNewFiles() {
  const validation = validateConfiguration();
  if (!validation.valid) {
    Logger.log('Configuration incomplete, skipping file check');
    return;
  }

  const newFiles = checkForNewAudioFiles();
  Logger.log(`Found ${newFiles.length} new audio files`);

  for (const fileInfo of newFiles) {
    const discussionId = processNewAudioFile(fileInfo);
    Logger.log(`Created discussion: ${discussionId}`);

    try {
      // Synchronous transcription with ElevenLabs
      startTranscription(discussionId);

      // Run speaker identification and populate SpeakerMap
      processSpeakerMapping(discussionId);
    } catch (e) {
      const timestamp = new Date().toISOString();
      Logger.log(`Failed to process ${discussionId}: ${e.message}`);
      updateDiscussion(discussionId, {
        status: CONFIG.STATUS.ERROR,
        error_message: `[${timestamp}] Transcription/mapping: ${e.message}`
      });
    }
  }
}

// ============================================================================
// SPEAKER MAPPING
// ============================================================================

/**
 * Run Gemini speaker identification and populate the SpeakerMap sheet.
 * Also generates named transcript, extracts teacher feedback, and creates summary.
 *
 * @param {string} discussionId
 */
function processSpeakerMapping(discussionId) {
  Logger.log(`Processing speaker mapping for ${discussionId}`);

  const transcript = getTranscript(discussionId);
  if (!transcript || !transcript.raw_transcript) {
    throw new Error('No transcript found for speaker mapping');
  }

  const rawTranscript = String(transcript.raw_transcript);

  // Get first ~3 minutes of transcript for speaker ID
  const lines = rawTranscript.split('\n');
  const excerptLines = [];
  for (const line of lines) {
    excerptLines.push(line);
    // Check if we've reached ~3 minutes by parsing timestamps
    const timeMatch = line.match(/\[(\d+):/);
    if (timeMatch && parseInt(timeMatch[1]) >= 3) break;
  }
  const excerpt = excerptLines.join('\n') || rawTranscript.substring(0, 5000);

  // Gemini speaker identification
  const speakerMap = identifySpeakers(excerpt);

  // Also get all speaker labels from transcript (in case Gemini missed some)
  const allLabels = extractSpeakerLabels(rawTranscript);

  // Populate SpeakerMap sheet
  for (const label of allLabels) {
    const suggestedName = speakerMap[label] || '?';
    upsertSpeakerMapping(discussionId, label, suggestedName);
  }

  // Build map from SpeakerMap sheet and generate named transcript
  const confirmedMap = buildSpeakerMapObject(discussionId);
  const namedTranscript = applySpeakerNames(rawTranscript, confirmedMap);

  // Save to transcript and discussion records
  upsertTranscript(discussionId, {
    speaker_map: JSON.stringify(confirmedMap),
    named_transcript: namedTranscript
  });

  updateDiscussion(discussionId, {
    status: CONFIG.STATUS.MAPPING,
    next_step: isGroupMode()
      ? 'Enter grade on this row, then click "Generate Feedback"'
      : 'Review speaker map in SpeakerMap sheet, confirm all speakers, then click "Generate Feedback"'
  });

  Logger.log(`Speaker mapping complete for ${discussionId}`);
}

// ============================================================================
// STATUS ADVANCEMENT
// ============================================================================

/**
 * Auto-advance discussions from mapping → review where appropriate.
 * In group mode: auto-advance since speaker map is auto-confirmed.
 * In individual mode: only advance if teacher has confirmed all speakers.
 */
function advanceMappingStatus() {
  const mappingDiscussions = getDiscussionsByStatus(CONFIG.STATUS.MAPPING);

  for (const discussion of mappingDiscussions) {
    const discussionId = discussion.discussion_id;

    if (isGroupMode()) {
      // Group mode: auto-advance to review
      updateDiscussion(discussionId, {
        status: CONFIG.STATUS.REVIEW,
        next_step: 'Enter grade on this row, then click "Generate Feedback"'
      });
      Logger.log(`Auto-advanced ${discussionId} to review (group mode)`);
    } else if (isSpeakerMapConfirmed(discussionId)) {
      // Individual mode: advance only when speakers confirmed
      createStudentReportsFromSpeakerMap(discussionId);
      updateDiscussion(discussionId, {
        status: CONFIG.STATUS.REVIEW,
        next_step: 'Enter grades in StudentReports, then click "Generate Feedback"'
      });
      Logger.log(`Advanced ${discussionId} to review (speakers confirmed)`);
    }
  }
}

/**
 * Create StudentReport rows from confirmed SpeakerMap entries (individual mode).
 * @param {string} discussionId
 */
function createStudentReportsFromSpeakerMap(discussionId) {
  const discussion = getDiscussion(discussionId);
  const transcript = getTranscript(discussionId);
  const namedTranscript = String(transcript.named_transcript || '');
  const speakerMap = buildSpeakerMapObject(discussionId);
  const studentNames = getStudentNames(speakerMap);

  for (const studentName of studentNames) {
    // Find or create student record
    let student = getStudentByName(studentName, discussion.section);
    if (!student) {
      const studentId = upsertStudent({
        name: studentName,
        email: '',
        section: discussion.section || ''
      });
      student = { student_id: studentId };
    }

    // Skip if report already exists
    const existing = getStudentReport(discussionId, student.student_id);
    if (existing) continue;

    // Extract this student's contributions
    const contributions = extractStudentContributions(namedTranscript, studentName);

    createStudentReport({
      discussion_id: discussionId,
      student_id: student.student_id,
      student_name: studentName,
      transcript_contributions: contributions
    });
  }

  Logger.log(`Created student reports for ${studentNames.length} students`);
}

// ============================================================================
// FEEDBACK GENERATION (menu action)
// ============================================================================

/**
 * Generate feedback for discussions in review status.
 * Group mode: generates group_feedback on the Discussion row.
 * Individual mode: generates per-student feedback on StudentReport rows.
 */
function generateFeedbackForDiscussions() {
  const discussions = [
    ...getDiscussionsByStatus(CONFIG.STATUS.REVIEW),
    ...getDiscussionsByStatus(CONFIG.STATUS.MAPPING)
  ];

  let generated = 0;

  for (const discussion of discussions) {
    const discussionId = discussion.discussion_id;

    try {
      const transcript = getTranscript(discussionId);
      if (!transcript) continue;

      const namedTranscript = String(transcript.named_transcript || transcript.raw_transcript || '');

      if (isGroupMode()) {
        // Group mode: need grade on Discussion row
        const grade = discussion.grade;
        if (!grade) {
          updateDiscussion(discussionId, {
            next_step: 'Enter grade on this row, then click "Generate Feedback" again'
          });
          continue;
        }

        const feedback = generateGroupFeedback(namedTranscript, String(grade));

        updateDiscussion(discussionId, {
          group_feedback: feedback,
          status: CONFIG.STATUS.REVIEW,
          next_step: 'Review feedback, check approved, then click "Send Approved Feedback"'
        });

        generated++;
      } else {
        // Individual mode: need per-student grades
        if (!isSpeakerMapConfirmed(discussionId)) {
          updateDiscussion(discussionId, {
            next_step: 'Confirm all speakers in SpeakerMap sheet first'
          });
          continue;
        }

        // Ensure StudentReports exist
        const reports = getReportsForDiscussion(discussionId);
        if (reports.length === 0) {
          createStudentReportsFromSpeakerMap(discussionId);
        }

        const updatedReports = getReportsForDiscussion(discussionId);
        let studentsNeedGrades = false;

        for (const report of updatedReports) {
          if (!report.grade) {
            studentsNeedGrades = true;
            continue;
          }

          // Skip if feedback already generated
          if (report.feedback) continue;

          try {
            Utilities.sleep(500);
            const contributions = String(report.transcript_contributions || '');
            const feedback = generateIndividualFeedback(
              report.student_name,
              contributions,
              namedTranscript,
              String(report.grade)
            );

            updateStudentReport(report.report_id, { feedback: feedback });
            generated++;
          } catch (e) {
            Logger.log(`Error generating feedback for ${report.student_name}: ${e.message}`);
            updateStudentReport(report.report_id, {
              feedback: `ERROR: ${e.message}`
            });
          }
        }

        if (studentsNeedGrades) {
          updateDiscussion(discussionId, {
            next_step: 'Enter grades for all students in StudentReports, then click "Generate Feedback" again'
          });
        } else {
          updateDiscussion(discussionId, {
            status: CONFIG.STATUS.REVIEW,
            next_step: 'Review feedback, approve students, then click "Send Approved Feedback"'
          });
        }
      }
    } catch (e) {
      const timestamp = new Date().toISOString();
      Logger.log(`Error generating feedback for ${discussionId}: ${e.message}`);
      updateDiscussion(discussionId, {
        status: CONFIG.STATUS.ERROR,
        error_message: `[${timestamp}] Feedback generation: ${e.message}`
      });
    }
  }

  return generated;
}

// ============================================================================
// SEND APPROVED FEEDBACK (menu action)
// ============================================================================

/**
 * Send feedback for approved discussions via enabled channels.
 * Checks distribute_email and distribute_canvas settings.
 */
function sendApprovedFeedback() {
  const emailEnabled = getSetting('distribute_email') === 'true';
  const canvasEnabled = getSetting('distribute_canvas') === 'true' && isCanvasConfigured();

  const discussions = [
    ...getDiscussionsByStatus(CONFIG.STATUS.REVIEW),
    ...getDiscussionsByStatus(CONFIG.STATUS.APPROVED)
  ];

  let totalEmailsSent = 0;
  let totalCanvasPosted = 0;
  let totalFailed = 0;

  for (const discussion of discussions) {
    const discussionId = discussion.discussion_id;

    // Check if approved
    if (isGroupMode()) {
      if (discussion.approved !== true) continue;
    } else {
      const reports = getApprovedUnsentReports(discussionId);
      if (reports.length === 0) continue;
    }

    // Send emails
    let discussionFailed = 0;
    const discussionErrors = [];

    if (emailEnabled) {
      try {
        const emailResult = sendAllReportsForDiscussion(discussionId);
        totalEmailsSent += emailResult.sent;
        discussionFailed += emailResult.failed;
        discussionErrors.push(...emailResult.errors);
      } catch (e) {
        Logger.log(`Email error for ${discussionId}: ${e.message}`);
        discussionFailed++;
        discussionErrors.push(`Email error: ${e.message}`);
      }
    }

    // Post to Canvas
    if (canvasEnabled && discussion.canvas_assignment_id) {
      try {
        const canvasResult = postGradesForDiscussion(discussionId);
        totalCanvasPosted += canvasResult.success;
        discussionFailed += canvasResult.failed;
        discussionErrors.push(...canvasResult.errors);
      } catch (e) {
        Logger.log(`Canvas error for ${discussionId}: ${e.message}`);
        discussionFailed++;
        discussionErrors.push(`Canvas error: ${e.message}`);
      }
    }

    totalFailed += discussionFailed;

    // Only mark 'sent' if all deliveries succeeded
    if (discussionFailed === 0) {
      updateDiscussion(discussionId, {
        status: CONFIG.STATUS.SENT,
        next_step: 'Feedback sent!'
      });
    } else {
      const timestamp = new Date().toISOString();
      const existingErrors = discussion.error_message || '';
      const newError = `[${timestamp}] Partial send failure (${discussionFailed} failed): ${discussionErrors.join('; ')}`;
      updateDiscussion(discussionId, {
        next_step: `Send incomplete — ${discussionFailed} failed. Re-run Send to retry.`,
        error_message: existingErrors ? existingErrors + '\n' + newError : newError
      });
    }
  }

  return {
    emailsSent: totalEmailsSent,
    canvasPosted: totalCanvasPosted,
    failed: totalFailed,
    emailEnabled: emailEnabled,
    canvasEnabled: canvasEnabled
  };
}

// ============================================================================
// FAILSAFES
// ============================================================================

/**
 * Detect discussions stuck in transcribing status for too long.
 * Marks them as error with a helpful message.
 */
function checkStuckTranscriptions() {
  const transcribing = getDiscussionsByStatus(CONFIG.STATUS.TRANSCRIBING);
  const now = Date.now();

  for (const discussion of transcribing) {
    const updatedAt = new Date(discussion.updated_at).getTime();
    const elapsed = now - updatedAt;

    if (elapsed > CONFIG.LIMITS.TRANSCRIPTION_TIMEOUT_MS) {
      const timestamp = new Date().toISOString();
      Logger.log(`Discussion ${discussion.discussion_id} stuck in transcribing for ${Math.round(elapsed / 60000)} minutes`);
      updateDiscussion(discussion.discussion_id, {
        status: CONFIG.STATUS.ERROR,
        error_message: `[${timestamp}] Transcription timed out after ${Math.round(elapsed / 60000)} minutes. Try splitting the audio file into shorter segments.`
      });
    }
  }
}

// ============================================================================
// CANVAS CONFIGURATION
// ============================================================================

/**
 * 3-step prompt dialog for Canvas course configuration.
 * Stores credentials, sets course ID, and auto-syncs all rosters.
 */
function showCanvasConfigDialog() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties().getProperties();

  // Step 1: Canvas base URL
  const existingUrl = props['CANVAS_BASE_URL'] || getSetting('canvas_base_url') || '';
  const urlResponse = ui.prompt('Configure Canvas Course (Step 1 of 3)',
    'Enter your Canvas base URL:\n(e.g., https://yourschool.instructure.com)' +
    (existingUrl ? '\n\nCurrent: ' + existingUrl : ''),
    ui.ButtonSet.OK_CANCEL);
  if (urlResponse.getSelectedButton() !== ui.Button.OK) return;
  const baseUrl = urlResponse.getResponseText().trim().replace(/\/$/, '') || existingUrl;
  if (!baseUrl) {
    ui.alert('Error', 'Canvas base URL is required.', ui.ButtonSet.OK);
    return;
  }

  // Step 2: API token
  const existingToken = props['CANVAS_API_TOKEN'] ? '(already set)' : '';
  const tokenResponse = ui.prompt('Configure Canvas Course (Step 2 of 3)',
    'Enter your Canvas API token:\n(Generate at Canvas → Account → Settings → New Access Token)' +
    (existingToken ? '\n\nLeave blank to keep existing token.' : ''),
    ui.ButtonSet.OK_CANCEL);
  if (tokenResponse.getSelectedButton() !== ui.Button.OK) return;
  const token = tokenResponse.getResponseText().trim();
  if (!token && !props['CANVAS_API_TOKEN']) {
    ui.alert('Error', 'Canvas API token is required.', ui.ButtonSet.OK);
    return;
  }

  // Step 3: Course ID
  const existingCourseId = getSetting('canvas_course_id') || '';
  const courseResponse = ui.prompt('Configure Canvas Course (Step 3 of 3)',
    'Enter the Canvas course ID:\n(Find it in the URL: canvas.school.edu/courses/12345 → 12345)' +
    (existingCourseId ? '\n\nCurrent: ' + existingCourseId : ''),
    ui.ButtonSet.OK_CANCEL);
  if (courseResponse.getSelectedButton() !== ui.Button.OK) return;
  const courseId = courseResponse.getResponseText().trim() || existingCourseId;
  if (!courseId) {
    ui.alert('Error', 'Course ID is required.', ui.ButtonSet.OK);
    return;
  }

  // Store configuration
  setProperty('CANVAS_BASE_URL', baseUrl);
  if (token) {
    setProperty('CANVAS_API_TOKEN', token);
  }
  setSetting('canvas_base_url', baseUrl);
  setSetting('canvas_course_id', courseId);
  setSetting('distribute_canvas', 'true');

  // Auto-sync all rosters
  try {
    const sections = getCanvasSections(courseId);
    const sectionMap = {};
    for (const section of sections) {
      sectionMap[section.id] = section.name;
    }

    const fallbackSection = sections.length === 1 ? sections[0].name : '';
    const count = syncCanvasStudents(courseId, fallbackSection, sectionMap);

    ui.alert('Canvas Configured',
      `Canvas course configured and roster synced!\n\n` +
      `Synced ${count} students from ${sections.length} section(s).\n` +
      `Sections: ${sections.map(s => s.name).join(', ')}\n\n` +
      `To switch courses or re-sync, run this dialog again.`,
      ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Partial Success',
      'Canvas credentials saved, but roster sync failed:\n' + e.message +
      '\n\nCheck your API token and course ID, then try again.',
      ui.ButtonSet.OK);
  }
}

// ============================================================================
// MENU ACTION WRAPPERS (with UI feedback)
// ============================================================================

function menuGenerateFeedback() {
  const ui = SpreadsheetApp.getUi();
  try {
    const count = generateFeedbackForDiscussions();
    ui.alert('Generate Feedback',
      `Generated feedback for ${count} item(s).\n\nCheck the Discussions and StudentReports sheets.`,
      ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', e.message, ui.ButtonSet.OK);
  }
}

function menuSendApprovedFeedback() {
  const ui = SpreadsheetApp.getUi();
  try {
    const result = sendApprovedFeedback();
    let message = '';
    if (result.emailEnabled) {
      message += `Emails sent: ${result.emailsSent}\n`;
    } else {
      message += 'Email distribution: disabled in Settings\n';
    }
    if (result.canvasEnabled) {
      message += `Canvas grades posted: ${result.canvasPosted}\n`;
    } else {
      message += 'Canvas distribution: disabled in Settings\n';
    }
    if (result.failed > 0) {
      message += `\nFailed: ${result.failed} (check Execution log for details)`;
    }
    ui.alert('Send Feedback', message, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', e.message, ui.ButtonSet.OK);
  }
}
