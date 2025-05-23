# Changelog

All notable changes to **Canvas Quiz Creator for Google Sheets** will be documented in this file. 
*(Adjust "Canvas Quiz Creator for Google Sheets" if you have a different preferred name for this specific tool)*

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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
