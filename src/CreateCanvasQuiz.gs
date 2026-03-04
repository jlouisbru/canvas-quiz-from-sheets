// ========= GLOBAL CONSTANTS ==========================================
const SHEET_SELECT_QUESTIONS_CONFIG = "SelectQuestions_Config";
const SHEET_SELECTED_QUESTIONS = "Selected Questions";
// const SHEET_DOC_EXAM_CONFIG = "DocExam_Config"; // REMOVED
const SHEET_CANVAS_QUIZ_CONFIG = "CanvasQuiz_Config";

// --- Column Header Constants for SelectQuestions_Config ---
const COL_SQ_CONFIG_SHEET_NAME = "Sheet Name";
const COL_SQ_CONFIG_INCLUDE = "Include";
const COL_SQ_CONFIG_TOTAL_QUESTIONS_FROM_SHEET = "Total Questions from Sheet";
const COL_SQ_CONFIG_QUESTIONS_PER_LECTURE = "Questions per Lecture";
const COL_SQ_CONFIG_LECTURE_RANGE_MIN = "Lecture Range Min";
const COL_SQ_CONFIG_LECTURE_RANGE_MAX = "Lecture Range Max";
const COL_SQ_CONFIG_DO_NOT_PICK_AGAIN = "Do Not Pick Again";

// --- Common Column Header Constants for Question Sheets ---
const COL_DO_NOT_PICK = "Do not pick";
const COL_LECTURE = "Lecture";
const COL_ID = "ID";
const COL_QUESTION_TEXT = "Question Text";
const COL_CORRECT_ANSWER = "Correct Answer";
const COL_OPTION_A = "Option A";
const COL_OPTION_B = "Option B";
const COL_OPTION_C = "Option C";
const COL_OPTION_D = "Option D";
const COL_OPTION_E = "Option E"; 
const COL_OPTION_F = "Option F"; 

// --- Column Header Constants for CreateQuiz_Config / CanvasQuiz_Config ---
const COL_CQ_CONFIG_PARAMETER = "Parameter";
const COL_CQ_CONFIG_VALUE = "Value";

// --- Script Property Keys ---
const PROP_CANVAS_API_TOKEN = 'CANVAS_API_TOKEN';

// ========= SPREADSHEET UI MENU =======================================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Quiz Tools')
    .addItem('Select Questions', 'runShuffleQuestions')
    .addSeparator()
    .addItem('Create Quiz on Canvas', 'runQuizCreation')
    .addToUi();
}

// ========= QUESTION SHUFFLER MODULE ==================================
var QuestionShuffler = {
  getConfig: function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName(SHEET_SELECT_QUESTIONS_CONFIG);
    if (!configSheet) throw new Error(`Config sheet '${SHEET_SELECT_QUESTIONS_CONFIG}' not found!`);

    const data = configSheet.getDataRange().getValues();
    const HEADER_ROW_IN_SHEET = 2, DATA_START_ROW_IN_SHEET = 3;
    const HEADER_ROW_INDEX = HEADER_ROW_IN_SHEET - 1, DATA_START_ROW_INDEX = DATA_START_ROW_IN_SHEET - 1;

    if (data.length <= HEADER_ROW_INDEX) throw new Error(`'${SHEET_SELECT_QUESTIONS_CONFIG}': Insufficient rows for headers (expected in row ${HEADER_ROW_IN_SHEET}).`);
    
    const configHeaders = data[HEADER_ROW_INDEX].map(h => String(h).trim());
    const headerMap = configHeaders.reduce((acc, h, i) => { acc[h] = i; return acc; }, {});
    const requiredHeaders = [COL_SQ_CONFIG_SHEET_NAME, COL_SQ_CONFIG_INCLUDE, COL_SQ_CONFIG_TOTAL_QUESTIONS_FROM_SHEET, COL_SQ_CONFIG_QUESTIONS_PER_LECTURE, COL_SQ_CONFIG_LECTURE_RANGE_MIN, COL_SQ_CONFIG_LECTURE_RANGE_MAX, COL_SQ_CONFIG_DO_NOT_PICK_AGAIN];
    const missing = requiredHeaders.filter(h => headerMap[h] === undefined);
    if (missing.length > 0) throw new Error(`'${SHEET_SELECT_QUESTIONS_CONFIG}': Missing columns (expected in row ${HEADER_ROW_IN_SHEET}): ${missing.join(', ')}.`);

    const config = { sheets: {} };
    for (let i = DATA_START_ROW_INDEX; i < data.length; i++) {
      const row = data[i];
      const sheetName = String(row[headerMap[COL_SQ_CONFIG_SHEET_NAME]]).trim();
      if (!sheetName) {
        if (row.some(cell => String(cell).trim())) Logger.log(`'${SHEET_SELECT_QUESTIONS_CONFIG}' (Row ${i + 1}): Skipping due to blank Sheet Name.`);
        continue;
      }
      const lectureMinRaw = row[headerMap[COL_SQ_CONFIG_LECTURE_RANGE_MIN]], lectureMaxRaw = row[headerMap[COL_SQ_CONFIG_LECTURE_RANGE_MAX]];
      let minLec = parseFloat(lectureMinRaw), maxLec = parseFloat(lectureMaxRaw);
      minLec = (lectureMinRaw === undefined || String(lectureMinRaw).trim() === "" || isNaN(minLec)) ? -Infinity : minLec;
      maxLec = (lectureMaxRaw === undefined || String(lectureMaxRaw).trim() === "" || isNaN(maxLec)) ? Infinity : maxLec;
       if ((lectureMinRaw !== undefined && String(lectureMinRaw).trim() !== "" && minLec === -Infinity && lectureMinRaw != '-Infinity') || 
          (lectureMaxRaw !== undefined && String(lectureMaxRaw).trim() !== "" && maxLec === Infinity && lectureMaxRaw != 'Infinity')) {
          Logger.log(`'${SHEET_SELECT_QUESTIONS_CONFIG}' (Row ${i+1}, Sheet '${sheetName}'): Invalid non-numeric lecture range. Skipping.`); continue;
      }
      if (minLec > maxLec) {
          Logger.log(`'${SHEET_SELECT_QUESTIONS_CONFIG}' (Row ${i+1}, Sheet '${sheetName}'): Lecture Min > Max. Skipping.`); continue;
      }
      config.sheets[sheetName] = {
        include: row[headerMap[COL_SQ_CONFIG_INCLUDE]] === true,
        questionsToSelect: parseInt(row[headerMap[COL_SQ_CONFIG_TOTAL_QUESTIONS_FROM_SHEET]]) || 0,
        questionsPerLecture: parseInt(row[headerMap[COL_SQ_CONFIG_QUESTIONS_PER_LECTURE]]) || 0,
        lectureRange: { min: minLec, max: maxLec },
        doNotPickAgain: row[headerMap[COL_SQ_CONFIG_DO_NOT_PICK_AGAIN]] === true
      };
    }
    if (Object.keys(config.sheets).length === 0 && data.length > DATA_START_ROW_INDEX) Logger.log(`'${SHEET_SELECT_QUESTIONS_CONFIG}': No valid sheet configurations parsed.`);
    return config;
  },

  shuffleQuestions: function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = this.getConfig();

    let targetSheet = ss.getSheetByName(SHEET_SELECTED_QUESTIONS);
    if (!targetSheet) targetSheet = ss.insertSheet(SHEET_SELECTED_QUESTIONS);
    else targetSheet.clear();

    const sourceSheetsData = {}; 
    for (const name in config.sheets) {
      if (config.sheets[name].include) {
        const sheet = ss.getSheetByName(name);
        if (sheet) sourceSheetsData[name] = sheet;
        else Logger.log(`Source sheet '${name}' not found.`);
      }
    }

    let allSelectedData = [];
    let finalTargetHeaders = null;
    const availableSheetNames = Object.keys(sourceSheetsData);

    for (const sheetName of availableSheetNames) {
      const sourceSheet = sourceSheetsData[sheetName];
      const sheetConfig = config.sheets[sheetName];
      
      const data = sourceSheet.getDataRange().getValues();
      if (data.length === 0) {
        Logger.log(`Sheet '${sheetName}' has no data rows.`);
        continue;
      }

      const currentSourceHeaders = data.shift().map(h => String(h).trim());
      if (finalTargetHeaders === null && currentSourceHeaders.length > 0) {
        finalTargetHeaders = [...currentSourceHeaders];
      }
      if (data.length === 0) {
        Logger.log(`Sheet '${sheetName}' has no data rows.`);
        continue;
      }

      const idIndex = currentSourceHeaders.indexOf(COL_ID);
      const lectureIndex = currentSourceHeaders.indexOf(COL_LECTURE);
      const doNotPickIndex = currentSourceHeaders.indexOf(COL_DO_NOT_PICK);
      if ([idIndex, lectureIndex, doNotPickIndex].some(idx => idx === -1) || (sheetConfig.doNotPickAgain && idIndex === -1)) {
        throw new Error(`Sheet '${sheetName}': Missing required columns (ID, Lecture, Do not pick). 'ID' also needed if 'Do Not Pick Again' is true.`);
      }
      
      const filtered = data.filter(row => {
          if (row[doNotPickIndex] === true) return false;
          const lecValRaw = row[lectureIndex];
          const lecVal = parseFloat(String(lecValRaw).trim());
          if (lecValRaw === undefined || String(lecValRaw).trim() === "") return sheetConfig.lectureRange.min === -Infinity && sheetConfig.lectureRange.max === Infinity;
          if (isNaN(lecVal)) {
              Logger.log(`Sheet '${sheetName}', ID '${row[idIndex]}': Non-numeric lecture '${lecValRaw}'. Skipping for range filter.`);
              return false;
          }
          return lecVal >= sheetConfig.lectureRange.min && lecVal <= sheetConfig.lectureRange.max;
      });

      const grouped = this.groupByLecture(filtered, lectureIndex);
      const selected = this.selectQuestionsPerLecture(grouped, sheetConfig.questionsPerLecture, sheetConfig.questionsToSelect);
      allSelectedData.push(...selected);

      if (sheetConfig.doNotPickAgain && selected.length > 0) {
        this.updateSourceSheet(sourceSheet, selected, doNotPickIndex, idIndex);
      }
    }

    if (!finalTargetHeaders || finalTargetHeaders.length === 0) { 
        targetSheet.getRange(1, 1).setValue(allSelectedData.length > 0 ? "Error: Questions selected, but headers undetermined." : "No source sheets processed or no headers found.");
        Logger.log("CRITICAL: finalTargetHeaders is null or empty. Cannot proceed.");
        return;
    }

    if (allSelectedData.length === 0) {
      targetSheet.getRange(1, 1).setValue("No questions found matching criteria.");
      targetSheet.getRange(1, 1, 1, finalTargetHeaders.length).setValues([finalTargetHeaders]); 
      Logger.log("Wrote headers, but no questions selected.");
      return;
    }

    const outputRows = allSelectedData.map(row => 
        finalTargetHeaders.map((_, i) => (row[i] !== undefined ? row[i] : ""))
    );

    targetSheet.getRange(1, 1, 1, finalTargetHeaders.length).setValues([finalTargetHeaders]); 
    targetSheet.getRange(2, 1, outputRows.length, finalTargetHeaders.length).setValues(outputRows); 
    targetSheet.autoResizeColumns(1, finalTargetHeaders.length);
    Logger.log("Successfully wrote headers and selected questions.");

    // The block for deleting "Do not pick" column was previously removed by user.
    // If it were to be re-added, it would go here, guarded by finalTargetHeaders check.
  },

  updateSourceSheet: function(sheet, selectedRows, dnpIdx, idIdx) {
    const numRows = sheet.getLastRow() - 1;
    if (numRows <= 0) return;
    const idColData  = sheet.getRange(2, idIdx  + 1, numRows, 1).getValues();
    const dnpColData = sheet.getRange(2, dnpIdx + 1, numRows, 1).getValues();
    const selectedIds = new Set(selectedRows.map(r => (r.length > idIdx && r[idIdx] !== null) ? String(r[idIdx]).trim() : null).filter(id => id));

    if (selectedIds.size === 0) {
      Logger.log(`'${sheet.getName()}': No valid IDs in selected questions for 'Do Not Pick Again'.`);
      return;
    }
    let changed = false;
    for (let i = 0; i < numRows; i++) {
      if (idColData[i][0] !== null) {
        const currentId = String(idColData[i][0]).trim();
        if (currentId && selectedIds.has(currentId) && dnpColData[i][0] !== true) {
          dnpColData[i][0] = true;
          changed = true;
        }
      }
    }
    if (changed) sheet.getRange(2, dnpIdx + 1, numRows, 1).setValues(dnpColData);
  },

  groupByLecture: function(data, lecIdx) {
    return data.reduce((acc, r) => {
      if (r.length > lecIdx) { const l = String(r[lecIdx]).trim(); (acc[l] = acc[l] || []).push(r); }
      return acc;
    }, {});
  },

  selectQuestionsPerLecture: function(grouped, perLec, total) {
    // perLec = 0 means no per-lecture limit: all questions go to rem and are randomly drawn up to total.
    let sel = [], rem = [];
    Object.values(grouped).forEach(qs => {
      this.shuffleArray(qs);
      const take = (perLec > 0) ? Math.min(perLec, qs.length) : 0;
      sel.push(...qs.slice(0, take));
      rem.push(...qs.slice(take));
    });
    this.shuffleArray(rem);
    if (sel.length < total) sel.push(...rem.slice(0, Math.max(0, total - sel.length)));
    return sel.slice(0, total);
  },

  shuffleArray: function(arr) {
    for (let i = arr.length - 1; i > 0; i--) { const j = Math.floor(Math.random()*(i+1)); [arr[i],arr[j]]=[arr[j],arr[i]]; }
  }
};

function runShuffleQuestions() {
  try {
    QuestionShuffler.shuffleQuestions();
    SpreadsheetApp.getUi().alert('Success', `Questions selected and placed in "${SHEET_SELECTED_QUESTIONS}" sheet.`, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error in Select Questions', e.message, SpreadsheetApp.getUi().ButtonSet.OK);
    Logger.log(`Error in runShuffleQuestions: ${e.toString()}\n${e.stack}`);
  }
}

// ========= DOC EXAM CREATOR MODULE (REVISED) ========= // REMOVED ENTIRE MODULE
// All functions related to Doc Exam Creator were here.

// ========= CANVAS QUIZ CREATOR MODULE =========

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
 * Returns { canvasQuestionType, questionPoints, answers } on success.
 * Returns { skip: true, error: '...' } if the question must be skipped.
 * Returns { ..., errorNote: '...' } for non-fatal issues (e.g. invalid TF answer).
 * The caller is responsible for pushing error/errorNote into questionsWithErrors.
 */
function resolveQuestion({ row, cols, questionId, correctAnswerFromSheet, config, rowNum, qText }) {
  if (questionId.startsWith("ES")) {
    return { canvasQuestionType: 'essay_question', questionPoints: config.POINTS_PER_QUESTION_ES, answers: [] };
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
    return { canvasQuestionType: 'true_false_question', questionPoints: config.POINTS_PER_QUESTION_TF, answers, errorNote };
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
    return { skip: true, error: `Sheet Row ${rowNum + 1} ("${qText.substring(0, 15)}..."): Skipped - No options for ${canvasQuestionType}.` };
  }
  if (isMC && questionPoints > 0 && answers.every(ans => ans.weight === 0) && correctAnswers.length > 0 && correctAnswers[0] !== "") {
    Logger.log(`Warning MC/MA Q (Row ${rowNum + 1}, "${qText.substring(0, 20)}..."): Correct "${correctAnswers.join(',')}" specified but no option matched.`);
  }

  return { canvasQuestionType, questionPoints, answers };
}

function getCanvasQuizModuleConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(SHEET_CANVAS_QUIZ_CONFIG);
  if (!configSheet) throw new Error(`Config sheet "${SHEET_CANVAS_QUIZ_CONFIG}" not found.`);

  const configData = configSheet.getDataRange().getValues();
  const HEADER_ROW_INDEX_IN_DATA = 1; 
  const DATA_START_ROW_INDEX_IN_DATA = 2; 

  if (configData.length < DATA_START_ROW_INDEX_IN_DATA + 1) {
      throw new Error(`'${SHEET_CANVAS_QUIZ_CONFIG}' sheet has insufficient rows.`);
  }

  const headers = configData[HEADER_ROW_INDEX_IN_DATA].map(h => String(h).trim());
  const paramCol = headers.indexOf(COL_CQ_CONFIG_PARAMETER);
  const valueCol = headers.indexOf(COL_CQ_CONFIG_VALUE);

  if (paramCol === -1 || valueCol === -1) {
      throw new Error(`In "${SHEET_CANVAS_QUIZ_CONFIG}", cols "${COL_CQ_CONFIG_PARAMETER}" & "${COL_CQ_CONFIG_VALUE}" not found in header row.`);
  }

  const config = {};
  for (let i = DATA_START_ROW_INDEX_IN_DATA; i < configData.length; i++) {
    const row = configData[i];
    if (!row || row.length === 0 || (row[paramCol] === undefined && row[valueCol] === undefined) ) continue;
    if (row.length <= Math.max(paramCol, valueCol) && (row[paramCol] || "").toString().trim() !== "") {
        Logger.log(`Warning: Skipping row ${i+1} in '${SHEET_CANVAS_QUIZ_CONFIG}' (param "${row[paramCol]}").`);
        continue;
    }
    const parameter = row[paramCol]?.toString().trim();
    let value = row[valueCol];

    if (parameter) {
      if (typeof value === 'string') {
        const lowerVal = value.toLowerCase();
        if (lowerVal === 'true') value = true;
        else if (lowerVal === 'false') value = false;
        // Numeric parameters (POINTS_*) are explicitly converted via parseFloat below.
        // No auto-cast here to avoid silently coercing string IDs or titles that happen to be numeric.
      }
      config[parameter] = value;
    }
  }

  const scriptProps = PropertiesService.getScriptProperties();
  let apiToken = scriptProps.getProperty(PROP_CANVAS_API_TOKEN);
  if (!apiToken || apiToken.trim() === '') {
      if (config.CANVAS_API_TOKEN && String(config.CANVAS_API_TOKEN).trim() !== '') {
          apiToken = String(config.CANVAS_API_TOKEN).trim();
          scriptProps.setProperty(PROP_CANVAS_API_TOKEN, apiToken);
      } else {
          const ui = SpreadsheetApp.getUi();
          const r = ui.prompt('Canvas API Token Required', 'Enter Canvas API Token:', ui.ButtonSet.OK_CANCEL);
          if (r.getSelectedButton() === ui.Button.OK && r.getResponseText().trim() !== '') {
              apiToken = r.getResponseText().trim();
              scriptProps.setProperty(PROP_CANVAS_API_TOKEN, apiToken);
          } else throw new Error('Canvas API Token not provided or prompt canceled.');
      }
  }
  config.CANVAS_API_TOKEN = apiToken;
  
  const baseReqP = ['CANVAS_API_URL', 'CANVAS_API_TOKEN', 'COURSE_ID', 'SHEET_NAME', 'QUIZ_TITLE'];
  const pointParams = ['POINTS_PER_QUESTION_TF', 'POINTS_PER_QUESTION_MC', 'POINTS_PER_QUESTION_ES']; 
  const optionalPointDefault = 'POINTS_PER_QUESTION_DEFAULT';

  const missingP = baseReqP.filter(p => !(p in config) || config[p] == null || String(config[p]).trim() === "");
  if (missingP.length > 0) throw new Error(`Missing required config params in "${SHEET_CANVAS_QUIZ_CONFIG}": ${missingP.join(', ')}.`);

  pointParams.forEach(pKey => {
    if (!(pKey in config) || config[pKey] == null || String(config[pKey]).trim() === "") {
      throw new Error(`Missing required point parameter "${pKey}" in "${SHEET_CANVAS_QUIZ_CONFIG}".`);
    }
    if (isNaN(parseFloat(config[pKey])) || parseFloat(config[pKey]) < 0) {
      throw new Error(`Parameter "${pKey}" in "${SHEET_CANVAS_QUIZ_CONFIG}" must be a non-negative number. Found: ${config[pKey]}`);
    }
    config[pKey] = parseFloat(config[pKey]); 
  });

  if (config[optionalPointDefault] !== undefined && config[optionalPointDefault] !== null && String(config[optionalPointDefault]).trim() !== "") {
      if (isNaN(parseFloat(config[optionalPointDefault])) || parseFloat(config[optionalPointDefault]) < 0) {
          throw new Error(`Optional parameter "${optionalPointDefault}" in "${SHEET_CANVAS_QUIZ_CONFIG}" must be a non-negative number if provided. Found: ${config[optionalPointDefault]}`);
      }
      config[optionalPointDefault] = parseFloat(config[optionalPointDefault]);
  } else {
      if (!(optionalPointDefault in config) || config[optionalPointDefault] == null || String(config[optionalPointDefault]).trim() === "") {
          Logger.log(`"${optionalPointDefault}" not found or is blank in config. Defaulting to 1 point for unmatched question types.`);
          config[optionalPointDefault] = 1;
      }
  }

  const validQuizTypes = ['assignment', 'practice_quiz', 'graded_survey', 'survey'];
  const quizTypeRaw = config.QUIZ_TYPE ? String(config.QUIZ_TYPE).trim().toLowerCase() : '';
  if (!quizTypeRaw) {
    config.QUIZ_TYPE = 'assignment';
  } else if (!validQuizTypes.includes(quizTypeRaw)) {
    throw new Error(`Invalid QUIZ_TYPE "${config.QUIZ_TYPE}" in "${SHEET_CANVAS_QUIZ_CONFIG}". Must be one of: ${validQuizTypes.join(', ')}.`);
  } else {
    config.QUIZ_TYPE = quizTypeRaw;
  }

  return config;
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

  // --- Create the quiz shell ---
  const authHeader = { 'Authorization': `Bearer ${config.CANVAS_API_TOKEN}` };
  const quizUrl = `${config.CANVAS_API_URL}/api/v1/courses/${config.COURSE_ID}/quizzes`;
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
  const quizCanvasId = quiz.id;
  const quizEditUrl = `${config.CANVAS_API_URL}/courses/${config.COURSE_ID}/quizzes/${quizCanvasId}/edit`;

  // --- Add questions ---
  let questionSheetRowNum = 0, successfulAdds = 0, currentQuizTotalPoints = 0;
  const questionsWithErrors = [];
  const totalQuestionsToProcess = questionRows.length;
  const qAddUrl = `${config.CANVAS_API_URL}/api/v1/courses/${config.COURSE_ID}/quizzes/${quizCanvasId}/questions`;

  for (const row of questionRows) {
    questionSheetRowNum++;
    if (questionSheetRowNum % 5 === 0 || questionSheetRowNum === totalQuestionsToProcess) {
      showProgressToast(`Processing question ${questionSheetRowNum} of ${totalQuestionsToProcess} for "${config.QUIZ_TITLE}"...`, 'Canvas Quiz Progress', 3);
    }

    const qText = row[cols.q]?.toString().trim();
    if (!qText) {
      questionsWithErrors.push(`Sheet Row ${questionSheetRowNum + 1}: Skipped - Empty question text.`);
      continue;
    }

    const questionId = row[cols.id]?.toString().trim().toUpperCase() || "";
    const rawCorrectAnswer = row[cols.correct];
    const correctAnswerFromSheet = typeof rawCorrectAnswer === 'boolean'
      ? String(rawCorrectAnswer).toUpperCase()
      : String(rawCorrectAnswer || "").trim().toUpperCase();

    const resolved = resolveQuestion({ row, cols, questionId, correctAnswerFromSheet, config, rowNum: questionSheetRowNum, qText });
    if (resolved.skip) { questionsWithErrors.push(resolved.error); continue; }
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
  }
  // --- Update quiz total points ---
  if (quiz.points_possible !== currentQuizTotalPoints) {
    Logger.log(`Updating quiz total points from ${quiz.points_possible} to ${currentQuizTotalPoints}.`);
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

  // --- Show summary ---
  const summaryItems = [
    `Canvas Quiz Creation for "${config.QUIZ_TITLE}" (ID: ${quizCanvasId}) Processed.`,
    `- Total questions from sheet: ${totalQuestionsToProcess}`,
    `- Successfully added to Canvas: ${successfulAdds}`
  ];
  if (questionsWithErrors.length > 0) {
    summaryItems.push(`- Questions with issues: ${questionsWithErrors.length}`, "\nDetails of issues (first few):");
    questionsWithErrors.slice(0, 5).forEach(err => summaryItems.push(`  - ${err}`));
    if (questionsWithErrors.length > 5) summaryItems.push("  - ... (see logs for all errors)");
  }
  summaryItems.push(
    `\nThe quiz is UNPUBLISHED. Final points: ${currentQuizTotalPoints}.`,
    "\nTo edit the quiz, copy and paste this URL:",
    quizEditUrl,
    "\nCheck Script Logs for full details."
  );
  showStandardDialog('Canvas Quiz Creation Summary', buildSummaryMessage(summaryItems));
  Logger.log(`Canvas quiz creation: ${successfulAdds} of ${questionRows.length} rows processed. ${successfulAdds} questions added to quiz ID ${quizCanvasId}. Total points: ${currentQuizTotalPoints}.`);
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
  try {
    showProgressToast('Starting Canvas Quiz setup validation...', 'Canvas Quiz');
    const { valid, messages, config, sheet } = validateQuizSetup();
    if (messages.length > 0) {
      showStandardDialog('Canvas Quiz Setup Validation', buildSummaryMessage(messages));
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
