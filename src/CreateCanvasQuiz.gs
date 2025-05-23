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
  SpreadsheetApp.getUi().createMenu('Extra Menu')
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
    if (availableSheetNames.length > 0) {
        for (const sheetName of availableSheetNames) {
            const sheet = sourceSheetsData[sheetName];
            const range = sheet.getDataRange();
            if (range) {
                const values = range.getValues();
                if (values.length > 0 && values[0].length > 0) { 
                    finalTargetHeaders = values[0].map(h => String(h).trim());
                    break; 
                }
            }
        }
    }
    
    for (const sheetName of availableSheetNames) { 
      const sourceSheet = sourceSheetsData[sheetName];
      const sheetConfig = config.sheets[sheetName];
      
      const data = sourceSheet.getDataRange().getValues();
      if (data.length <= 1) { 
        Logger.log(`Sheet '${sheetName}' has no data rows.`);
        continue; 
      }

      const currentSourceHeaders = data.shift().map(h => String(h).trim());
      if (finalTargetHeaders === null && currentSourceHeaders.length > 0) {
          finalTargetHeaders = [...currentSourceHeaders];
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
    const data = sheet.getRange(2, 1, numRows, sheet.getLastColumn()).getValues();
    const dnpColData = sheet.getRange(2, dnpIdx + 1, numRows, 1).getValues();
    const selectedIds = new Set(selectedRows.map(r => (r.length > idIdx && r[idIdx] !== null) ? String(r[idIdx]).trim() : null).filter(id => id));

    if (selectedIds.size === 0) {
      Logger.log(`'${sheet.getName()}': No valid IDs in selected questions for 'Do Not Pick Again'.`);
      return;
    }
    let changed = false;
    for (let i = 0; i < numRows; i++) {
      if (data[i].length > idIdx && data[i][idIdx] !== null) {
        const currentId = String(data[i][idIdx]).trim();
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
        else if (value.trim() !== '' && !isNaN(value) &&
                 parameter !== 'COURSE_ID' && 
                 !parameter.endsWith('_URL') && 
                 !parameter.endsWith('_TOKEN') && 
                 parameter !== PROP_CANVAS_API_TOKEN &&
                 !parameter.startsWith('QUIZ_TITLE') && 
                 !parameter.startsWith('SHEET_NAME') 
                 ) { 
          value = Number(value);
        }
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
  return config;
}

function createQuizInCanvas(config, sheet) {
  const quizData = sheet.getDataRange().getValues(), questionRows = quizData.slice(1);
  const headers = quizData[0].map(h => String(h).trim());
  const cols = {
      q: headers.indexOf(COL_QUESTION_TEXT), 
      correct: headers.indexOf(COL_CORRECT_ANSWER),
      id: headers.indexOf(COL_ID),
      a: headers.indexOf(COL_OPTION_A), b: headers.indexOf(COL_OPTION_B), 
      c: headers.indexOf(COL_OPTION_C), d: headers.indexOf(COL_OPTION_D),
      e: headers.indexOf(COL_OPTION_E), f: headers.indexOf(COL_OPTION_F)
  };

  if ([cols.q, cols.correct, cols.id].some(idx => idx === -1)) {
      throw new Error(`Sheet "${config.SHEET_NAME}" missing required columns (ID, Question Text, Correct Answer).`);
  }
  
  const placeholderTotalPts = 0; 
  const quizUrl = `${config.CANVAS_API_URL}/api/v1/courses/${config.COURSE_ID}/quizzes`;
  const quizPayloadShell = { 
      'quiz[title]': config.QUIZ_TITLE, 
      'quiz[quiz_type]': 'assignment', 
      'quiz[points_possible]': placeholderTotalPts, 
      'quiz[published]': false 
  };
  const quizOpts = {method: 'post', headers: {'Authorization': `Bearer ${config.CANVAS_API_TOKEN}`}, payload: quizPayloadShell, muteHttpExceptions: true};
  const quizResp = UrlFetchApp.fetch(quizUrl, quizOpts);

  if (quizResp.getResponseCode() >= 300) {
    const errorMsg = `Failed to create quiz shell for "${config.QUIZ_TITLE}".\nStatus: ${quizResp.getResponseCode()}.\nResponse: ${quizResp.getContentText().substring(0,200)}`;
    // showStandardDialog('Canvas Quiz Creation Error', errorMsg); // Assuming UiUtils.gs
    SpreadsheetApp.getUi().alert('Canvas Quiz Creation Error', errorMsg, SpreadsheetApp.getUi().ButtonSet.OK);
    throw new Error(errorMsg.replace(/\n/g, ' '));
  }
  const quiz = JSON.parse(quizResp.getContentText());
  const quizCanvasId = quiz.id;
  const quizEditUrl = `${config.CANVAS_API_URL}/courses/${config.COURSE_ID}/quizzes/${quizCanvasId}/edit`;

  let questionSheetRowNum = 0, successfulAdds = 0;
  let currentQuizTotalPoints = 0;
  const questionsWithErrors = [];
  const totalQuestionsToProcess = questionRows.length;

  for (const row of questionRows) {
    questionSheetRowNum++;
    // if (questionSheetRowNum % 5 === 0 || questionSheetRowNum === totalQuestionsToProcess) {
    //     showProgressToast(`Processing question ${questionSheetRowNum} of ${totalQuestionsToProcess} for "${config.QUIZ_TITLE}"...`, 'Canvas Quiz Progress', 3); // Assuming UiUtils.gs
    // }

    const qText = row[cols.q]?.toString().trim();
    if (!qText) {
        questionsWithErrors.push(`Sheet Row ${questionSheetRowNum+1}: Skipped - Empty question text.`);
        continue;
    }

    const questionId = row[cols.id]?.toString().trim().toUpperCase() || "";
    const rawCorrectAnswer = row[cols.correct];
    
    let correctAnswerFromSheet;
    if (typeof rawCorrectAnswer === 'boolean') {
        correctAnswerFromSheet = String(rawCorrectAnswer).toUpperCase(); 
    } else {
        correctAnswerFromSheet = String(rawCorrectAnswer || "").trim().toUpperCase();
    }
    
    let questionPoints = 0;
    let canvasQuestionType = ''; 
    let answersPayloadForCanvas = [];

    if (questionId.startsWith("ES")) {
        canvasQuestionType = 'essay_question';
        questionPoints = config.POINTS_PER_QUESTION_ES;
    } else if (questionId.startsWith("TF")) {
        canvasQuestionType = 'true_false_question';
        questionPoints = config.POINTS_PER_QUESTION_TF;
        if (correctAnswerFromSheet === "TRUE") {
            answersPayloadForCanvas.push({ text: "True",  weight: 100 });
            answersPayloadForCanvas.push({ text: "False", weight: 0 });
        } else if (correctAnswerFromSheet === "FALSE") {
            answersPayloadForCanvas.push({ text: "True",  weight: 0 });
            answersPayloadForCanvas.push({ text: "False", weight: 100 });
        } else {
            Logger.log(`ERROR for TF Question ID "${questionId}" (Row ${questionSheetRowNum+1}): Correct answer is "${correctAnswerFromSheet}", which is not "TRUE" or "FALSE".`);
            questionsWithErrors.push(`Sheet Row ${questionSheetRowNum+1} (TF "${qText.substring(0,15)}..."): Invalid correct answer specified: '${correctAnswerFromSheet}'.`);
            answersPayloadForCanvas.push({ text: "True",  weight: 0 }); 
            answersPayloadForCanvas.push({ text: "False", weight: 0 });
        }
    } else if (questionId.startsWith("MC")) {
        const mcCorrectAnswers = correctAnswerFromSheet.split(',').map(l => l.trim()).filter(l => l);
        if (mcCorrectAnswers.length > 1) canvasQuestionType = 'multiple_answers_question';
        else canvasQuestionType = 'multiple_choice_question';
        questionPoints = config.POINTS_PER_QUESTION_MC;
        const optionSlots = [ { letter: 'A', index: cols.a }, { letter: 'B', index: cols.b }, { letter: 'C', index: cols.c }, { letter: 'D', index: cols.d }, { letter: 'E', index: cols.e }, { letter: 'F', index: cols.f }];
        optionSlots.forEach(slot => {
          if (slot.index !== -1 && row.length > slot.index && row[slot.index] != null && String(row[slot.index]).trim() !== "") {
            const optionText = String(row[slot.index]).trim();
            let weight = 0;
            if (canvasQuestionType === 'multiple_answers_question') {
                if (mcCorrectAnswers.includes(slot.letter)) weight = 100 / Math.max(1, mcCorrectAnswers.length);
            } else { 
                if (mcCorrectAnswers.includes(slot.letter)) weight = 100;
            }
            answersPayloadForCanvas.push({text: optionText, weight: weight});
          }
        });
         if (answersPayloadForCanvas.length === 0) { 
            Logger.log(`Skipping MC/MA Q (Row ${questionSheetRowNum+1}, "${qText.substring(0,20)}..."): No valid options.`); 
            questionsWithErrors.push(`Sheet Row ${questionSheetRowNum+1} ("${qText.substring(0,15)}..."): Skipped - No options for MC/MA.`);
            continue; 
        }
        if (questionPoints > 0 && answersPayloadForCanvas.every(ans => ans.weight === 0) && mcCorrectAnswers.length > 0 && mcCorrectAnswers[0] !== "") {
            Logger.log(`Warning MC/MA Q (Row ${questionSheetRowNum+1}, "${qText.substring(0,20)}..."): Correct "${mcCorrectAnswers.join(',')}" specified, but no option had weight.`);
        }
    } else { 
        const defaultCorrectAnswers = correctAnswerFromSheet.split(',').map(l => l.trim()).filter(l => l);
        if (defaultCorrectAnswers.length > 1) canvasQuestionType = 'multiple_answers_question';
        else canvasQuestionType = 'multiple_choice_question';
        questionPoints = config.POINTS_PER_QUESTION_DEFAULT;
        Logger.log(`ID "${questionId}" (Row ${questionSheetRowNum+1}) not TF/MC/ES. Type: ${canvasQuestionType}, Points: ${questionPoints}`);
        const optionSlots = [ { letter: 'A', index: cols.a }, { letter: 'B', index: cols.b }, { letter: 'C', index: cols.c }, { letter: 'D', index: cols.d }, { letter: 'E', index: cols.e }, { letter: 'F', index: cols.f }];
        optionSlots.forEach(slot => {
          if (slot.index !== -1 && row.length > slot.index && row[slot.index] != null && String(row[slot.index]).trim() !== "") {
            const optionText = String(row[slot.index]).trim();
            let weight = 0;
            if (canvasQuestionType === 'multiple_answers_question') {
                if (defaultCorrectAnswers.includes(slot.letter)) weight = 100 / Math.max(1, defaultCorrectAnswers.length);
            } else { 
                if (defaultCorrectAnswers.includes(slot.letter)) weight = 100;
            }
            answersPayloadForCanvas.push({text: optionText, weight: weight});
          }
        });
         if (answersPayloadForCanvas.length === 0 && canvasQuestionType !== 'essay_question') { 
            Logger.log(`Skipping Default Q (Row ${questionSheetRowNum+1}, "${qText.substring(0,20)}..."): No valid options.`); 
            questionsWithErrors.push(`Sheet Row ${questionSheetRowNum+1} ("${qText.substring(0,15)}..."): Skipped - No options for Default type.`);
            continue; 
        }
    }

    if (questionPoints === undefined || questionPoints === null) {
        Logger.log(`Error: Points undefined for ID "${questionId}" (Row ${questionSheetRowNum+1}). Skipping.`);
        questionsWithErrors.push(`Sheet Row ${questionSheetRowNum+1} ("${qText.substring(0,15)}..."): Points undefined for type.`);
        continue;
    }
    
    const questionAPIPayload = {
        'question[question_name]': `Question ${successfulAdds + 1}`, 
        'question[question_text]': qText, 
        'question[question_type]': canvasQuestionType, 
        'question[points_possible]': questionPoints
    };

    if (answersPayloadForCanvas.length > 0) { 
        answersPayloadForCanvas.forEach((ans, idx) => { 
            questionAPIPayload[`question[answers][${idx}][answer_text]`] = ans.text; 
            questionAPIPayload[`question[answers][${idx}][answer_weight]`] = ans.weight;
        });
    } else if (canvasQuestionType !== 'essay_question') {
        Logger.log(`Error Q (Row ${questionSheetRowNum+1}, "${qText.substring(0,20)}..."): Type ${canvasQuestionType} but answersPayloadForCanvas empty. Skipping.`);
        questionsWithErrors.push(`Sheet Row ${questionSheetRowNum+1} ("${qText.substring(0,15)}..."): Error - ${canvasQuestionType} empty answers payload.`);
        continue;
    }
    
    const qAddUrl = `${config.CANVAS_API_URL}/api/v1/courses/${config.COURSE_ID}/quizzes/${quizCanvasId}/questions`;
    const qAddOpts = {method:'post', headers:{'Authorization':`Bearer ${config.CANVAS_API_TOKEN}`}, payload:questionAPIPayload, muteHttpExceptions:true};
    const qAddResp = UrlFetchApp.fetch(qAddUrl, qAddOpts); 
    
    if (qAddResp.getResponseCode()>=300) {
        const errorDetail = `Failed to add Q (Row ${questionSheetRowNum+1}, "${qText.substring(0,15)}...", Type: ${canvasQuestionType}). Status ${qAddResp.getResponseCode()}.`;
        Logger.log(`${errorDetail} Canvas Response: ${qAddResp.getContentText().substring(0,300)} Payload Sent (partial): ${JSON.stringify(questionAPIPayload).substring(0,300)}`);
        questionsWithErrors.push(errorDetail);
    } else {
        successfulAdds++;
        currentQuizTotalPoints += questionPoints;
        const addedQuestionCanvasId = JSON.parse(qAddResp.getContentText()).id;
        if ((successfulAdds) !== parseInt(questionAPIPayload['question[question_name]'].split(' ')[1])) {
            const updateQNamePayload = {'question[question_name]': `Question ${successfulAdds}`};
            const qUpdateUrl = `${config.CANVAS_API_URL}/api/v1/courses/${config.COURSE_ID}/quizzes/${quizCanvasId}/questions/${addedQuestionCanvasId}`;
            UrlFetchApp.fetch(qUpdateUrl, {method:'put', headers:{'Authorization':`Bearer ${config.CANVAS_API_TOKEN}`}, payload:updateQNamePayload, muteHttpExceptions:true});
        }
    }
    SpreadsheetApp.flush();
  }

  if (quiz.points_possible !== currentQuizTotalPoints) { 
    Logger.log(`Quiz shell points: ${quiz.points_possible}. Calculated successful adds points: ${currentQuizTotalPoints}. Updating.`);
    const updatePayload = {'quiz[points_possible]': currentQuizTotalPoints};
    const updateResp = UrlFetchApp.fetch(`${quizUrl}/${quizCanvasId}`,{method:'put',headers:{'Authorization':`Bearer ${config.CANVAS_API_TOKEN}`},payload:updatePayload,muteHttpExceptions:true});
    if (updateResp.getResponseCode() >= 300) Logger.log(`Failed to update quiz total points to ${currentQuizTotalPoints}. Status: ${updateResp.getResponseCode()}. Resp: ${updateResp.getContentText().substring(0,100)}`);
    else Logger.log(`Quiz total points updated to ${currentQuizTotalPoints}.`);
  }
  
  const summaryItems = [];
  summaryItems.push(`Canvas Quiz Creation for "${config.QUIZ_TITLE}" (ID: ${quizCanvasId}) Processed.`);
  summaryItems.push(`- Total questions from sheet: ${totalQuestionsToProcess}`);
  summaryItems.push(`- Successfully added to Canvas: ${successfulAdds}`);
  if (questionsWithErrors.length > 0) {
    summaryItems.push(`- Questions with issues: ${questionsWithErrors.length}`);
    summaryItems.push("\nDetails of issues (first few):");
    questionsWithErrors.slice(0, 5).forEach(err => summaryItems.push(`  - ${err}`));
    if (questionsWithErrors.length > 5) summaryItems.push("  - ... (see logs for all errors)");
  }
  summaryItems.push(`\nThe quiz is UNPUBLISHED. Final points: ${currentQuizTotalPoints}.`);
  summaryItems.push("\nTo edit the quiz, copy and paste this URL:");
  summaryItems.push(`${quizEditUrl}`);
  summaryItems.push("\nCheck Script Logs for full details.");

  // showStandardDialog('Canvas Quiz Creation Summary', buildSummaryMessage(summaryItems)); // Assuming UiUtils.gs
  SpreadsheetApp.getUi().alert('Canvas Quiz Creation Summary', summaryItems.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
  Logger.log(`Canvas quiz creation: ${successfulAdds} of ${questionRows.length} rows processed. ${successfulAdds} questions added to quiz ID ${quizCanvasId}. Total points: ${currentQuizTotalPoints}.`);
}

function runQuizCreation() {
  const ui = SpreadsheetApp.getUi(), ss = SpreadsheetApp.getActiveSpreadsheet(); 
  const valMsgs=[], addMsg = (msg, isErr = false) => { valMsgs.push(msg); if (isErr) valErrs = true; };
  let valErrs=false, config, qSheet; 

  try {
    // showProgressToast('Starting Canvas Quiz setup validation...', 'Canvas Quiz'); // Assuming UiUtils.gs
    try { 
      config = getCanvasQuizModuleConfig(); 
      addMsg('✅ CanvasQuiz_Config processed.'); 
      
      qSheet = ss.getSheetByName(config.SHEET_NAME);
      if (!qSheet) addMsg(`❌ Questions sheet "${config.SHEET_NAME}" not found.`, true);
      else {
        addMsg(`✅ Questions sheet "${config.SHEET_NAME}" found.`);
        const data = qSheet.getDataRange().getValues();
        if (data.length <= 1) addMsg(`❌ Questions sheet "${config.SHEET_NAME}" is empty or only has headers.`, true);
        else {
          addMsg('✅ Questions sheet has data.');
          const headers = data[0].map(h => String(h).trim());
          const reqH = [COL_QUESTION_TEXT, COL_CORRECT_ANSWER, COL_ID]; // Option columns checked within createQuizInCanvas for MC
          const missingH = reqH.filter(h => !headers.includes(h));
          if (missingH.length > 0) addMsg(`❌ Questions sheet "${config.SHEET_NAME}" missing: ${missingH.join(', ')}.`, true);
          else addMsg('✅ Critical columns present in questions sheet.');
        }
      }

      if (!valErrs && config.CANVAS_API_URL && config.CANVAS_API_TOKEN && config.COURSE_ID) { 
        const testUrl = `${config.CANVAS_API_URL}/api/v1/courses/${config.COURSE_ID}`;
        const testOptions = {method:'get', headers:{'Authorization':`Bearer ${config.CANVAS_API_TOKEN}`}, muteHttpExceptions:true};
        const testResp = UrlFetchApp.fetch(testUrl, testOptions);
        if (testResp.getResponseCode() === 200) addMsg('✅ Canvas API connection and Course ID seem valid.');
        else addMsg(`❌ Canvas API/Course ID invalid. Status: ${testResp.getResponseCode()}. Resp: ${testResp.getContentText().substring(0, 70)}...`, true);
      } else if (!valErrs) {
        addMsg('❌ Canvas API details incomplete; connection not tested.', true);
      }
    } catch (configValidationError) { 
      addMsg(`❌ CRITICAL SETUP ERROR: ${configValidationError.message}`, true); 
      Logger.log(`Validation/Config Error in runQuizCreation: ${configValidationError.stack}`); 
    }

    if (valMsgs.length > 0) {
      // showStandardDialog('Canvas Quiz Setup Validation', buildSummaryMessage(valMsgs)); // Assuming UiUtils.gs
      ui.alert('Canvas Quiz Setup Validation', valMsgs.join('\n'), ui.ButtonSet.OK);
    }
    if (valErrs) { Logger.log("Canvas Quiz creation aborted due to validation errors."); return; }

    const userConfirmation = ui.alert('Create Quiz on Canvas', `Proceed to create quiz "${config.QUIZ_TITLE}" on Canvas using questions from "${config.SHEET_NAME}"?\nQuiz will be UNPUBLISHED.`, ui.ButtonSet.YES_NO);

    if (userConfirmation === ui.Button.YES) {
      if (!qSheet) qSheet = ss.getSheetByName(config.SHEET_NAME); // Re-fetch just in case
      if (!qSheet) throw new Error("Question sheet became unavailable after validation.");
      // showProgressToast(`Initiating creation of quiz: "${config.QUIZ_TITLE}"...`, 'Canvas Quiz Creator', 7); // Assuming UiUtils.gs
      createQuizInCanvas(config, qSheet); 
    } else {
      // showStandardDialog('Cancelled', 'Canvas quiz creation was cancelled by the user.'); // Assuming UiUtils.gs
      ui.alert('Cancelled', 'Canvas quiz creation was cancelled.', ui.ButtonSet.OK);
    }
  } catch (runtimeError) { 
    // showStandardDialog('Error During Canvas Quiz Creation', `Runtime error: ${runtimeError.message}\nCheck script logs.`); // Assuming UiUtils.gs
    ui.alert('Error During Canvas Quiz Creation', `Runtime error: ${runtimeError.message}\nCheck script logs.`, ui.ButtonSet.OK);
    Logger.log(`Error in runQuizCreation (runtime): ${runtimeError.toString()}\n${runtimeError.stack}`);
  }
}
