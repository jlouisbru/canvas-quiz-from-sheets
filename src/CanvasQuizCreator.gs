// ========= CANVAS QUIZ CREATOR MODULE ====================================

/**
 * Wraps UrlFetchApp.fetch with automatic retry on rate-limit (429) or
 * temporary server errors (503). Waits 1 s, 2 s, 4 s between attempts.
 */
function canvasFetchWithRetry(url, opts, maxRetries = 3) {
  let resp;
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    resp = UrlFetchApp.fetch(url, opts);
    const code = resp.getResponseCode();
    if (code !== 429 && code !== 503) return resp;
    if (attempt < maxRetries) {
      const delay = Math.pow(2, attempt - 1) * 1000;
      Logger.log(`Canvas API rate limited (${code}), attempt ${attempt}/${maxRetries}. Retrying in ${delay}ms...`);
      Utilities.sleep(delay);
    }
  }
  return resp;
}

/**
 * Builds the answers payload array for a multiple-choice or multiple-answers question.
 * Returns an array of { text, weight } objects for any non-empty option slots.
 */
function buildMCAnswers(row, cols, correctAnswers, questionType) {
  const optionSlots = [
    { letter: 'A', index: cols.a }, { letter: 'B', index: cols.b },
    { letter: 'C', index: cols.c }, { letter: 'D', index: cols.d },
    { letter: 'E', index: cols.e }, { letter: 'F', index: cols.f }
  ];
  return optionSlots
    .filter(slot => slot.index !== -1 && row.length > slot.index && row[slot.index] != null && String(row[slot.index]).trim() !== "")
    .map(slot => {
      let weight = 0;
      if (questionType === 'multiple_answers_question') {
        if (correctAnswers.includes(slot.letter)) weight = 100 / Math.max(1, correctAnswers.length);
      } else {
        if (correctAnswers.includes(slot.letter)) weight = 100;
      }
      return { text: String(row[slot.index]).trim(), weight };
    });
}

/**
 * Determines the Canvas question type, point value, and answers for one row.
 * Returns one of two shapes (discriminated union):
 *   { ok: false, error: String }                              — fatal skip
 *   { ok: true, canvasQuestionType, questionPoints, answers,  — success
 *     errorNote?: String }                                       (errorNote is a non-fatal warning)
 * The caller is responsible for pushing error/errorNote into questionsWithErrors.
 */
function resolveQuestion({ row, cols, questionId, correctAnswerFromSheet, config, rowNum, qText }) {
  if (questionId.startsWith("ES")) {
    return { ok: true, canvasQuestionType: 'essay_question', questionPoints: config.POINTS_PER_QUESTION_ES, answers: [] };
  }

  if (questionId.startsWith("TF")) {
    let answers, errorNote = null;
    if (correctAnswerFromSheet === "TRUE") {
      answers = [{ text: "True", weight: 100 }, { text: "False", weight: 0 }];
    } else if (correctAnswerFromSheet === "FALSE") {
      answers = [{ text: "True", weight: 0 }, { text: "False", weight: 100 }];
    } else {
      Logger.log(`ERROR for TF Question ID "${questionId}" (Row ${rowNum + 1}): Correct answer is "${correctAnswerFromSheet}", expected "TRUE" or "FALSE".`);
      errorNote = `Sheet Row ${rowNum + 1} (TF "${qText.substring(0, 15)}..."): Invalid correct answer: '${correctAnswerFromSheet}'.`;
      answers = [{ text: "True", weight: 0 }, { text: "False", weight: 0 }];
    }
    return { ok: true, canvasQuestionType: 'true_false_question', questionPoints: config.POINTS_PER_QUESTION_TF, answers, errorNote };
  }

  const isMC = questionId.startsWith("MC");
  const correctAnswers = correctAnswerFromSheet.split(',').map(l => l.trim()).filter(l => l);
  const canvasQuestionType = correctAnswers.length > 1 ? 'multiple_answers_question' : 'multiple_choice_question';
  const questionPoints = isMC ? config.POINTS_PER_QUESTION_MC : config.POINTS_PER_QUESTION_DEFAULT;

  if (!isMC) {
    Logger.log(`ID "${questionId}" (Row ${rowNum + 1}) not TF/MC/ES. Treating as ${canvasQuestionType}, Points: ${questionPoints}`);
  }

  const answers = buildMCAnswers(row, cols, correctAnswers, canvasQuestionType);
  if (answers.length === 0) {
    Logger.log(`Skipping Q (Row ${rowNum + 1}, "${qText.substring(0, 20)}..."): No valid options found.`);
    return { ok: false, error: `Sheet Row ${rowNum + 1} ("${qText.substring(0, 15)}..."): Skipped - No options for ${canvasQuestionType}.` };
  }
  if (isMC && questionPoints > 0 && answers.every(ans => ans.weight === 0) && correctAnswers.length > 0 && correctAnswers[0] !== "") {
    Logger.log(`Warning MC/MA Q (Row ${rowNum + 1}, "${qText.substring(0, 20)}..."): Correct "${correctAnswers.join(',')}" specified but no option matched.`);
  }

  return { ok: true, canvasQuestionType, questionPoints, answers };
}

// ========= RESUME STATE HELPERS ==========================================
// These helpers checkpoint a running quiz creation to survive the Apps Script
// 6-minute execution-time limit. The constants are defined in Constants.gs.

function saveResumeState_(quizId, quizTitle, sheetName, lastRow, totalPoints) {
  const props = {};
  props[RESUME_PROP_QUIZ_ID]    = String(quizId);
  props[RESUME_PROP_QUIZ_TITLE] = quizTitle;
  props[RESUME_PROP_SHEET_NAME] = sheetName;
  props[RESUME_PROP_LAST_ROW]   = String(lastRow);
  props[RESUME_PROP_TOTAL_PTS]  = String(totalPoints);
  PropertiesService.getScriptProperties().setProperties(props);
}

function loadResumeState_() {
  const sp = PropertiesService.getScriptProperties();
  const quizId = sp.getProperty(RESUME_PROP_QUIZ_ID);
  if (!quizId) return null;
  return {
    quizId:      quizId,
    quizTitle:   sp.getProperty(RESUME_PROP_QUIZ_TITLE) || '',
    sheetName:   sp.getProperty(RESUME_PROP_SHEET_NAME) || '',
    lastRow:     parseInt(sp.getProperty(RESUME_PROP_LAST_ROW) || '-1', 10),
    totalPoints: parseFloat(sp.getProperty(RESUME_PROP_TOTAL_PTS) || '0')
  };
}

function clearResumeState_() {
  const sp = PropertiesService.getScriptProperties();
  [RESUME_PROP_QUIZ_ID, RESUME_PROP_QUIZ_TITLE, RESUME_PROP_SHEET_NAME,
   RESUME_PROP_LAST_ROW, RESUME_PROP_TOTAL_PTS].forEach(function(k) {
    sp.deleteProperty(k);
  });
}

// Returns true when a progress toast should be shown for the given 0-based question index.
// Shows every question for the first 10, then every 5 after that.
function shouldShowToast_(questionIndex) {
  if (questionIndex < 10) return true;
  return (questionIndex + 1) % 5 === 0;
}

function createQuizInCanvas(config, sheet) {
  const quizData = sheet.getDataRange().getValues();
  const questionRows = quizData.slice(1);
  const headers = quizData[0].map(h => String(h).trim());
  const cols = {
    q:       headers.indexOf(COL_QUESTION_TEXT),
    correct: headers.indexOf(COL_CORRECT_ANSWER),
    id:      headers.indexOf(COL_ID),
    a: headers.indexOf(COL_OPTION_A), b: headers.indexOf(COL_OPTION_B),
    c: headers.indexOf(COL_OPTION_C), d: headers.indexOf(COL_OPTION_D),
    e: headers.indexOf(COL_OPTION_E), f: headers.indexOf(COL_OPTION_F)
  };
  if ([cols.q, cols.correct, cols.id].some(idx => idx === -1)) {
    throw new Error(`Sheet "${config.SHEET_NAME}" missing required columns (ID, Question Text, Correct Answer).`);
  }

  const authHeader = { 'Authorization': `Bearer ${config.CANVAS_API_TOKEN}` };
  const quizUrl = `${config.CANVAS_API_URL}/api/v1/courses/${config.COURSE_ID}/quizzes`;

  // --- Resume support: skip shell creation if resuming a previous run ---
  const isResume = !!config._resumeQuizId;
  let quizCanvasId, quizInitialPoints;

  if (isResume) {
    quizCanvasId      = config._resumeQuizId;
    quizInitialPoints = config._resumeTotalPts;
    Logger.log(`Resuming quiz ID ${quizCanvasId} from row ${config._resumeLastRow + 1}.`);
  } else {
    const quizResp = canvasFetchWithRetry(quizUrl, {
      method: 'post',
      headers: authHeader,
      payload: {
        'quiz[title]': config.QUIZ_TITLE,
        'quiz[quiz_type]': config.QUIZ_TYPE,
        'quiz[points_possible]': 0,
        'quiz[published]': false
      },
      muteHttpExceptions: true
    });
    if (quizResp.getResponseCode() >= 300) {
      const errorMsg = `Failed to create quiz shell for "${config.QUIZ_TITLE}".\nStatus: ${quizResp.getResponseCode()}.\nResponse: ${quizResp.getContentText().substring(0, 200)}`;
      showStandardDialog('Canvas Quiz Creation Error', errorMsg);
      throw new Error(errorMsg.replace(/\n/g, ' '));
    }
    const quiz = JSON.parse(quizResp.getContentText());
    quizCanvasId      = quiz.id;
    quizInitialPoints = quiz.points_possible;
  }

  const quizEditUrl = `${config.CANVAS_API_URL}/courses/${config.COURSE_ID}/quizzes/${quizCanvasId}/edit`;
  const deleteUrl   = `${config.CANVAS_API_URL}/api/v1/courses/${config.COURSE_ID}/quizzes/${quizCanvasId}`;
  const qAddUrl     = `${config.CANVAS_API_URL}/api/v1/courses/${config.COURSE_ID}/quizzes/${quizCanvasId}/questions`;

  const resumeFromRow = isResume ? config._resumeLastRow + 1 : 0;
  let successfulAdds = 0, currentQuizTotalPoints = isResume ? config._resumeTotalPts : 0;
  const questionsWithErrors = [];
  const totalQuestionsToProcess = questionRows.length;

  try {
    // --- Add questions ---
    for (let i = resumeFromRow; i < questionRows.length; i++) {
      const row = questionRows[i];
      const questionSheetRowNum = i + 1; // 1-based for error messages
      if (shouldShowToast_(i) || i === totalQuestionsToProcess - 1) {
        const pct = Math.round(((i + 1) / totalQuestionsToProcess) * 100);
        showProgressToast(`Question ${i + 1} of ${totalQuestionsToProcess} (${pct}%) — "${config.QUIZ_TITLE}"`, 'Creating Quiz', 3);
      }

      const qText = row[cols.q]?.toString().trim();
      if (!qText) {
        questionsWithErrors.push(`Sheet Row ${questionSheetRowNum + 1}: Skipped - Empty question text.`);
        // Checkpoint even for skipped rows so resume starts after this row.
        saveResumeState_(quizCanvasId, config.QUIZ_TITLE, sheet.getName(), i, currentQuizTotalPoints);
        continue;
      }

      const questionId = row[cols.id]?.toString().trim().toUpperCase() || "";
      const rawCorrectAnswer = row[cols.correct];
      const correctAnswerFromSheet = typeof rawCorrectAnswer === 'boolean'
        ? String(rawCorrectAnswer).toUpperCase()
        : String(rawCorrectAnswer || "").trim().toUpperCase();

      const resolved = resolveQuestion({ row, cols, questionId, correctAnswerFromSheet, config, rowNum: questionSheetRowNum, qText });
      if (!resolved.ok) {
        questionsWithErrors.push(resolved.error);
        saveResumeState_(quizCanvasId, config.QUIZ_TITLE, sheet.getName(), i, currentQuizTotalPoints);
        continue;
      }
      if (resolved.errorNote) { questionsWithErrors.push(resolved.errorNote); }
      const { canvasQuestionType, questionPoints, answers } = resolved;
      const questionAPIPayload = {
        'question[question_name]': `Question ${successfulAdds + 1} (${questionId})`,
        'question[question_text]': qText,
        'question[question_type]': canvasQuestionType,
        'question[points_possible]': questionPoints
      };

      if (answers.length > 0) {
        answers.forEach((ans, idx) => {
          questionAPIPayload[`question[answers][${idx}][answer_text]`] = ans.text;
          questionAPIPayload[`question[answers][${idx}][answer_weight]`] = ans.weight;
        });
      } else if (canvasQuestionType !== 'essay_question') {
        Logger.log(`Error Q (Row ${questionSheetRowNum + 1}, "${qText.substring(0, 20)}..."): ${canvasQuestionType} has empty answers. Skipping.`);
        questionsWithErrors.push(`Sheet Row ${questionSheetRowNum + 1} ("${qText.substring(0, 15)}..."): Error - ${canvasQuestionType} empty answers payload.`);
        saveResumeState_(quizCanvasId, config.QUIZ_TITLE, sheet.getName(), i, currentQuizTotalPoints);
        continue;
      }

      const qAddResp = canvasFetchWithRetry(qAddUrl, { method: 'post', headers: authHeader, payload: questionAPIPayload, muteHttpExceptions: true });
      if (qAddResp.getResponseCode() >= 300) {
        const errorDetail = `Failed to add Q (Row ${questionSheetRowNum + 1}, "${qText.substring(0, 15)}...", Type: ${canvasQuestionType}). Status ${qAddResp.getResponseCode()}.`;
        Logger.log(`${errorDetail} Canvas Response: ${qAddResp.getContentText().substring(0, 300)} Payload (partial): ${JSON.stringify(questionAPIPayload).substring(0, 300)}`);
        questionsWithErrors.push(errorDetail);
      } else {
        successfulAdds++;
        currentQuizTotalPoints += questionPoints;
      }
      // Checkpoint after every row so a timeout can resume from here.
      saveResumeState_(quizCanvasId, config.QUIZ_TITLE, sheet.getName(), i, currentQuizTotalPoints);
    }

    // --- Update quiz total points ---
    if (quizInitialPoints !== currentQuizTotalPoints) {
      Logger.log(`Updating quiz total points from ${quizInitialPoints} to ${currentQuizTotalPoints}.`);
      const updateResp = canvasFetchWithRetry(`${quizUrl}/${quizCanvasId}`, {
        method: 'put',
        headers: authHeader,
        payload: { 'quiz[points_possible]': currentQuizTotalPoints },
        muteHttpExceptions: true
      });
      if (updateResp.getResponseCode() >= 300) {
        Logger.log(`Failed to update quiz total points. Status: ${updateResp.getResponseCode()}. Resp: ${updateResp.getContentText().substring(0, 100)}`);
      } else {
        Logger.log(`Quiz total points updated to ${currentQuizTotalPoints}.`);
      }
    }

    // Success — clear resume state and show summary.
    clearResumeState_();
    const summaryLines = [
      `Quiz: "${config.QUIZ_TITLE}" (Canvas ID: ${quizCanvasId})`,
      `Questions processed: ${totalQuestionsToProcess}`,
      `Successfully added: ${successfulAdds}`,
      `Total points: ${currentQuizTotalPoints}`,
      `Status: UNPUBLISHED`
    ];
    if (questionsWithErrors.length > 0) {
      summaryLines.push(`Issues (${questionsWithErrors.length}): see Script Logs for details.`);
    }
    showHtmlSummaryDialog('Canvas Quiz Creation Summary', summaryLines, quizEditUrl, 'Open Quiz Editor in Canvas');
    Logger.log(`Canvas quiz creation: ${successfulAdds} of ${questionRows.length} rows processed. ${successfulAdds} questions added to quiz ID ${quizCanvasId}. Total points: ${currentQuizTotalPoints}.`);

  } catch (fatalErr) {
    // A fatal runtime error occurred after the quiz shell was created.
    // Offer to delete the orphaned quiz so Canvas doesn't accumulate blank quizzes.
    Logger.log(`Fatal error during question add for quiz ${quizCanvasId}: ${fatalErr.toString()}\n${fatalErr.stack}`);
    clearResumeState_();
    const ui = SpreadsheetApp.getUi();
    const resp = ui.alert(
      'Quiz Creation Error',
      `A fatal error occurred after the quiz shell was created.\n\nError: ${fatalErr.message}\n\nDelete the incomplete quiz from Canvas?`,
      ui.ButtonSet.YES_NO
    );
    if (resp === ui.Button.YES) {
      try {
        canvasFetchWithRetry(deleteUrl, {
          method: 'delete',
          headers: authHeader,
          muteHttpExceptions: true
        });
        showStandardDialog('Cleanup Complete', 'The incomplete quiz has been deleted from Canvas.');
      } catch (deleteErr) {
        showHtmlSummaryDialog('Cleanup Failed',
          ['Could not auto-delete the quiz. Please delete it manually.'],
          quizEditUrl, 'Open Quiz in Canvas');
      }
    } else {
      showHtmlSummaryDialog('Incomplete Quiz Left in Canvas',
        ['The quiz was not deleted. You can edit or delete it manually.'],
        quizEditUrl, 'Open Quiz in Canvas');
    }
  }
}

function validateQuizSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const messages = [];
  let config, sheet, hasErrors = false;
  const addMsg = (msg, isErr = false) => { messages.push(msg); if (isErr) hasErrors = true; };

  try {
    config = getCanvasQuizModuleConfig();
    addMsg('✅ CanvasQuiz_Config processed.');

    sheet = ss.getSheetByName(config.SHEET_NAME);
    if (!sheet) {
      addMsg(`❌ Questions sheet "${config.SHEET_NAME}" not found.`, true);
    } else {
      addMsg(`✅ Questions sheet "${config.SHEET_NAME}" found.`);
      const data = sheet.getDataRange().getValues();
      if (data.length <= 1) {
        addMsg(`❌ Questions sheet "${config.SHEET_NAME}" is empty or only has headers.`, true);
      } else {
        addMsg('✅ Questions sheet has data.');
        const headers = data[0].map(h => String(h).trim());
        const missingH = [COL_QUESTION_TEXT, COL_CORRECT_ANSWER, COL_ID].filter(h => !headers.includes(h));
        if (missingH.length > 0) addMsg(`❌ Questions sheet "${config.SHEET_NAME}" missing: ${missingH.join(', ')}.`, true);
        else addMsg('✅ Critical columns present in questions sheet.');
      }
    }

    if (!hasErrors && config.CANVAS_API_URL && config.CANVAS_API_TOKEN && config.COURSE_ID) {
      const testResp = UrlFetchApp.fetch(
        `${config.CANVAS_API_URL}/api/v1/courses/${config.COURSE_ID}`,
        { method: 'get', headers: { 'Authorization': `Bearer ${config.CANVAS_API_TOKEN}` }, muteHttpExceptions: true }
      );
      if (testResp.getResponseCode() === 200) addMsg('✅ Canvas API connection and Course ID seem valid.');
      else addMsg(`❌ Canvas API/Course ID invalid. Status: ${testResp.getResponseCode()}. Resp: ${testResp.getContentText().substring(0, 70)}...`, true);
    } else if (!hasErrors) {
      addMsg('❌ Canvas API details incomplete; connection not tested.', true);
    }
  } catch (e) {
    addMsg(`❌ CRITICAL SETUP ERROR: ${e.message}`, true);
    Logger.log(`Validation/Config Error in validateQuizSetup: ${e.stack}`);
  }

  return { valid: !hasErrors, messages, config, sheet };
}

function runQuizCreation() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- Check for a previous run that was interrupted by execution-time limit ---
  const resume = loadResumeState_();
  if (resume) {
    const resumeResp = ui.alert(
      'Resume Previous Quiz Creation?',
      `A previous quiz creation was interrupted.\n\nQuiz: "${resume.quizTitle}"\nLast question processed: #${resume.lastRow + 1}\n\nYES to resume where it left off, NO to start fresh.`,
      ui.ButtonSet.YES_NO
    );
    if (resumeResp === ui.Button.YES) {
      const sheet = ss.getSheetByName(resume.sheetName);
      if (!sheet) {
        showStandardDialog('Resume Error', `Sheet "${resume.sheetName}" not found. Cannot resume. Starting fresh.`);
        clearResumeState_();
      } else {
        try {
          const config = getCanvasQuizModuleConfig();
          config._resumeQuizId   = resume.quizId;
          config._resumeLastRow  = resume.lastRow;
          config._resumeTotalPts = resume.totalPoints;
          showProgressToast(`Resuming quiz "${resume.quizTitle}" from question ${resume.lastRow + 2}...`, 'Canvas Quiz Creator', 7);
          createQuizInCanvas(config, sheet);
        } catch (resumeErr) {
          showStandardDialog('Error During Resume', `${resumeErr.message}\nCheck script logs.`);
          Logger.log(`Error resuming quiz creation: ${resumeErr.toString()}\n${resumeErr.stack}`);
        }
        return;
      }
    } else {
      clearResumeState_();
    }
  }

  // --- Normal creation flow ---
  try {
    showProgressToast('Starting Canvas Quiz setup validation...', 'Canvas Quiz');
    const { valid, messages, config, sheet } = validateQuizSetup();
    if (messages.length > 0) {
      showStandardDialog('Canvas Quiz Setup Validation', messages.join('\n'));
    }
    if (!valid) { Logger.log("Canvas Quiz creation aborted due to validation errors."); return; }

    const userConfirmation = ui.alert('Create Quiz on Canvas', `Proceed to create quiz "${config.QUIZ_TITLE}" on Canvas using questions from "${config.SHEET_NAME}"?\nQuiz will be UNPUBLISHED.`, ui.ButtonSet.YES_NO);
    if (userConfirmation === ui.Button.YES) {
      const qSheet = sheet || ss.getSheetByName(config.SHEET_NAME);
      if (!qSheet) throw new Error("Question sheet became unavailable after validation.");
      showProgressToast(`Initiating creation of quiz: "${config.QUIZ_TITLE}"...`, 'Canvas Quiz Creator', 7);
      createQuizInCanvas(config, qSheet);
    } else {
      showStandardDialog('Cancelled', 'Canvas quiz creation was cancelled by the user.');
    }
  } catch (runtimeError) {
    showStandardDialog('Error During Canvas Quiz Creation', `Runtime error: ${runtimeError.message}\nCheck script logs.`);
    Logger.log(`Error in runQuizCreation (runtime): ${runtimeError.toString()}\n${runtimeError.stack}`);
  }
}
