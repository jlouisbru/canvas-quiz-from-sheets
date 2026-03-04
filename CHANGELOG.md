# Changelog

All notable changes to **Canvas Quiz Creator for Google Sheets** will be documented in this file. 
*(Adjust "Canvas Quiz Creator for Google Sheets" if you have a different preferred name for this specific tool)*

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.1] - 2026-03-04

### Added

-   **Configurable Quiz Type:** New optional `QUIZ_TYPE` parameter in `CanvasQuiz_Config` supports all Canvas quiz types: `assignment` (default), `practice_quiz`, `graded_survey`, and `survey`. Invalid values are caught with a clear error message.
-   **API Rate Limit Handling:** Canvas API calls now automatically retry on 429 (rate limit) and 503 (server error) responses with exponential backoff (1 s, 2 s, 4 s), making large quiz uploads significantly more reliable.

### Improved

-   **Progress Toasts:** Progress toast notifications and summary dialogs are now properly displayed during quiz creation (previously defined but inactive).
-   **Code Refactoring:** Extracted `buildMCAnswers` and `resolveQuestion` helper functions to eliminate duplicated logic and improve readability.

### Fixed

-   **Performance:** `SpreadsheetApp.flush()` was being called on every question iteration; it is now called once after the loop, eliminating unnecessary sync overhead.
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
    -   Custom "Extra Menu" for easy operation.
    -   Secure API key handling with Script Properties (prompt on first use).
    -   Progress tracking with toast notifications and summary dialogs.
    -   Comprehensive "Instructions" sheet template provided for setup and usage guidance.
-   **Documentation:**
    -   `README.md` for project overview, setup, and usage.
    -   `LICENSE.md` (MIT).
    -   This `CHANGELOG.md`.
