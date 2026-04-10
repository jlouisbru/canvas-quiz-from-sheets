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
    // NOTE: dnpColData is intentionally mutated here for batch write performance.
    // It is a local 2-D array read from the sheet in this function scope, not shared state.
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
    /*
     * Two-pass selection strategy:
     * Pass 1 (guaranteed): for each lecture group, pick up to perLec questions at random.
     * Pass 2 (overflow): collect all remaining unselected questions into a pool;
     *   draw randomly to fill the gap between guaranteed count and the total target.
     * When perLec = 0, all questions go directly to overflow (no per-lecture cap).
     */
    let guaranteed = [], overflow = [];
    Object.values(grouped).forEach(qs => {
      const shuffled = this.shuffleArray(qs);
      const take = (perLec > 0) ? Math.min(perLec, shuffled.length) : 0;
      guaranteed.push(...shuffled.slice(0, take));
      overflow.push(...shuffled.slice(take));
    });
    overflow = this.shuffleArray(overflow);
    if (guaranteed.length < total) guaranteed.push(...overflow.slice(0, Math.max(0, total - guaranteed.length)));
    return guaranteed.slice(0, total);
  },

  shuffleArray: function(arr) {
    // Returns a NEW array; does not mutate the input.
    const copy = arr.slice();
    for (let i = copy.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [copy[i], copy[j]] = [copy[j], copy[i]];
    }
    return copy;
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
