// ========= CANVAS QUIZ CONFIG MODULE =====================================

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
  let tokenCellRange = null; // tracks the Range of the CANVAS_API_TOKEN value cell for masking
  for (let i = DATA_START_ROW_INDEX_IN_DATA; i < configData.length; i++) {
    const row = configData[i];
    if (!row || row.length === 0 || (row[paramCol] === undefined && row[valueCol] === undefined)) continue;
    if (row.length <= Math.max(paramCol, valueCol) && (row[paramCol] || "").toString().trim() !== "") {
      Logger.log(`Warning: Skipping row ${i + 1} in '${SHEET_CANVAS_QUIZ_CONFIG}' (param "${row[paramCol]}").`);
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
      if (parameter === 'CANVAS_API_TOKEN') {
        // Capture the Range now so we can mask it after saving to Script Properties.
        tokenCellRange = configSheet.getRange(i + 1, valueCol + 1);
      }
    }
  }

  const scriptProps = PropertiesService.getScriptProperties();
  let apiToken = scriptProps.getProperty(PROP_CANVAS_API_TOKEN);
  if (!apiToken || apiToken.trim() === '') {
    if (config.CANVAS_API_TOKEN && String(config.CANVAS_API_TOKEN).trim() !== '') {
      apiToken = String(config.CANVAS_API_TOKEN).trim();
      scriptProps.setProperty(PROP_CANVAS_API_TOKEN, apiToken);
      // Mask the token in the sheet so it is no longer stored in plain text.
      if (tokenCellRange) {
        tokenCellRange.setValue('•••••');
        SpreadsheetApp.flush();
      }
      showProgressToast('API token secured — moved to Script Properties', 'Security', 5);
    } else {
      const ui = SpreadsheetApp.getUi();
      const r = ui.prompt('Canvas API Token Required', 'Enter Canvas API Token:', ui.ButtonSet.OK_CANCEL);
      if (r.getSelectedButton() === ui.Button.OK && r.getResponseText().trim() !== '') {
        apiToken = r.getResponseText().trim();
        scriptProps.setProperty(PROP_CANVAS_API_TOKEN, apiToken);
      } else {
        throw new Error('Canvas API Token not provided or prompt canceled.');
      }
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
