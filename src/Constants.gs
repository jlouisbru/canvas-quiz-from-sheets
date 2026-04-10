// ========= GLOBAL CONSTANTS ==========================================

// --- Sheet names ---
const SHEET_SELECT_QUESTIONS_CONFIG = "SelectQuestions_Config";
const SHEET_SELECTED_QUESTIONS = "Selected Questions";
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

// --- Column Header Constants for CanvasQuiz_Config ---
const COL_CQ_CONFIG_PARAMETER = "Parameter";
const COL_CQ_CONFIG_VALUE = "Value";

// --- Script Property Keys ---
const PROP_CANVAS_API_TOKEN = 'CANVAS_API_TOKEN';

// --- Resume State Script Property Keys ---
// Used to checkpoint a running quiz creation so it can survive an
// Apps Script execution-time-limit timeout and be resumed.
const RESUME_PROP_QUIZ_ID    = 'RESUME_QUIZ_ID';
const RESUME_PROP_QUIZ_TITLE = 'RESUME_QUIZ_TITLE';
const RESUME_PROP_SHEET_NAME = 'RESUME_SHEET_NAME';
const RESUME_PROP_LAST_ROW   = 'RESUME_LAST_PROCESSED_ROW'; // 0-based index
const RESUME_PROP_TOTAL_PTS  = 'RESUME_TOTAL_POINTS';
