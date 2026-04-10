# Changelog

All notable changes to **Canvas Quiz Creator for Google Sheets** will be documented in this file. 
*(Adjust "Canvas Quiz Creator for Google Sheets" if you have a different preferred name for this specific tool)*

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.2] - 2026-04-10

### Added

-   **Execution Timeout Recovery:** Quiz creation now checkpoints progress after every question. If the Apps Script 6-minute execution limit is hit mid-run, re-running "Create Quiz on Canvas" prompts to resume from the last processed question rather than starting over.
-   **Orphan Quiz Cleanup:** If a fatal error occurs after the quiz shell has been created in Canvas, a dialog offers to automatically delete the incomplete quiz, preventing blank quizzes from accumulating in your course.
-   **Clickable Summary Dialog:** The quiz creation summary now shows a clickable "Open Quiz Editor in Canvas" link, replacing the plain-text alert.
-   **clasp Deployment Support:** A `.clasp.json` and `appsscript.json` are now included, enabling `clasp push` for direct deployment from the command line.

### Improved

-   **API Token Auto-Masking:** When a `CANVAS_API_TOKEN` is found in the `CanvasQuiz_Config` sheet and saved to Script Properties, the cell is immediately replaced with `•••••` so the token is never left in plain text.
-   **Adaptive Progress Toast:** Toast notifications now fire for every question during the first 10, then every 5 questions after — reducing notification noise on large uploads.
-   **Non-Mutating `shuffleArray`:** The array shuffle now returns a new array instead of mutating the input, eliminating a subtle side-effect risk.
-   **Clearer Selection Logic:** Variables `sel`/`rem` in `selectQuestionsPerLecture` renamed to `guaranteed`/`overflow`, with a block comment explaining the two-pass selection strategy.
-   **Code Split into Focused Modules:** `CreateCanvasQuiz.gs` has been broken into five purpose-specific files — `Constants.gs`, `Config.gs`, `QuestionShuffler.gs`, `CanvasQuizCreator.gs`, and `Main.gs` — each under 300 lines.

### Removed

-   **Dead Code:** `buildSummaryMessage` helper function removed (replaced by `showHtmlSummaryDialog`).

## [1.0.1] - 2026-03-04

### Added

-   **Configurable Quiz Type:** New optional `QUIZ_TYPE` parameter in `CanvasQuiz_Config` supports all Canvas quiz types: `assignment` (default), `practice_quiz`, `graded_survey`, and `survey`. Invalid values are caught with a clear error message.
-   **API Rate Limit Handling:** Canvas API calls now automatically retry on 429 (rate limit) and 503 (server error) responses with exponential backoff (1 s, 2 s, 4 s), making large quiz uploads significantly more reliable.
-   **Question ID in Canvas Question Name:** Canvas questions are now named `Question N (ID)` (e.g. `Question 1 (TF042)`) instead of just `Question N`, making it easy to trace which question bank item ended up in which quiz.

### Improved

-   **Progress Toasts:** Progress toast notifications and summary dialogs are now properly displayed during quiz creation (previously defined but inactive).
-   **Code Refactoring:** Extracted `buildMCAnswers` and `resolveQuestion` helper functions to eliminate duplicated logic and improve readability.
-   **Menu Renamed:** The custom spreadsheet menu is now labeled `Quiz Tools` instead of `Extra Menu`.
-   **`resolveQuestion` Refactored:** The function now takes a single destructured object instead of 8 positional parameters, and returns error information (`{ skip, error }` or `{ errorNote }`) instead of mutating the caller's error array — eliminating a hidden side effect.
-   **Validation Extracted:** `validateQuizSetup()` is now a standalone function, removing a nested try/catch from `runQuizCreation()` and making each concern independently readable and testable.
-   **Header Detection Simplified:** Removed a redundant pre-loop pass in `shuffleQuestions` that re-read every source sheet just to find headers. The main loop now handles this with a two-step empty check, capturing headers even from header-only sheets.
-   **`updateSourceSheet` Optimized:** Now reads only the ID and "Do not pick" columns instead of the full row data, reducing unnecessary API calls for wide sheets.
-   **`perLec = 0` Semantics Documented:** Added an inline comment clarifying that `questionsPerLecture = 0` means no per-lecture limit (all questions go into the random fill pool), not "take zero per lecture."

### Fixed

-   **Performance:** All `SpreadsheetApp.flush()` calls have been removed from the quiz creation flow — one was firing on every question iteration, and a second remained after the loop with no cells to flush. Neither was needed.
-   **Dead Code Removed:** Removed a question-name update block that made a redundant Canvas API PUT request on every successfully added question.
-   **Config Parsing:** Removed a fragile numeric auto-cast in config value parsing that could silently coerce string parameters (e.g., `COURSE_ID`, quiz titles) if they happened to look numeric. All numeric parameters (`POINTS_*`) are now explicitly converted.

## [1.0.0] - 2025-05-25

### Added

-   **Initial Release of Create Canvas Quiz from Google Sheets!**
-   **Question Selection System:**
    -   Select questions from multiple user-defined question bank sheets.
    -   Configuration via `SelectQuestions_Config` sheet (specify source sheets, count, lecture range, per-lecture limits).
    -   "Do Not Pick Again" feature to mark used questions in source banks.
    -   Selected questions compiled into a dedicated `Selected Questions` sheet.
-   **Canvas Quiz Uploader:**
    -   Create quizzes directly in Canvas LMS from the `Selected Questions` sheet.
    *   Configuration via `CanvasQuiz_Config` sheet (Canvas URL, Course ID, Quiz Title, API Token handling).
    -   Support for question-type specific point values (TF, MC, ES, Default) based on `ID` prefixes.
    -   Handles True/False, Multiple Choice, Multiple Answers, and Essay question types.
    -   Quizzes created as **UNPUBLISHED** in Canvas for review.
-   **Spreadsheet Interface & Setup:**
    -   Custom menu (originally "Extra Menu", renamed to "Quiz Tools" in v1.0.1) for easy operation.
    -   Secure API key handling with Script Properties (prompt on first use).
    -   Progress tracking with toast notifications and summary dialogs.
    -   Comprehensive "Instructions" sheet template provided for setup and usage guidance.
-   **Documentation:**
    -   `README.md` for project overview, setup, and usage.
    -   `LICENSE.md` (MIT).
    -   This `CHANGELOG.md`.
