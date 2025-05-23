# Create Canvas Quiz From Google Sheets

This Google Apps Script project streamlines the process of selecting questions from various banks within a Google Sheet, and then uploading those selected questions to create quizzes directly in Canvas LMS.

## ✨ Features

*   **Centralized Question Banks:** Manage True/False, Multiple Choice, and Essay questions in dedicated Google Sheet tabs.
*   **Flexible Question Selection:** Configure rules in `SelectQuestions_Config` to pick questions based on:
    *   Source sheet
    *   Total number of questions
    *   Questions per lecture/topic
    *   Lecture/topic range
    *   Option to mark questions as "Do Not Pick Again" in source sheets.
*   **Automated Canvas Quiz Creation:**
    *   Generates quizzes in Canvas using questions from the `Selected Questions` sheet.
    *   Supports different point values per question type (TF, MC, ES, Default).
    *   Handles True/False, Multiple Choice, Multiple Answers, and Essay question types in Canvas.
    *   Creates quizzes as **UNPUBLISHED**, allowing for review before student access.
*   **User-Friendly Interface:**
    *   Custom menu within Google Sheets for easy operation.
    *   Configuration driven through dedicated sheets (`SelectQuestions_Config`, `CanvasQuiz_Config`).
    *   Progress indicators and completion summaries.
*   **Instruction Sheet:** Includes a detailed "Instructions" tab within the spreadsheet for setup and usage guidance.

## 🚀 Getting Started

### Prerequisites

1.  A Google Account (to use Google Sheets and Apps Script).
2.  A Canvas LMS account with appropriate permissions to create quizzes and generate API Access Tokens.

### Setup Instructions

1.  **Make a Copy of the Google Sheet:**
    *   Open this [GOOGLE SHEET TEMPLATE (VIEW ONLY)](https://docs.google.com/spreadsheets/d/1mRXJ_Ei8BUdyw2S9E31uS_CtCJ58d1yxh3RnYdp6hQY/edit?usp=sharing).
    *   Go to **File > Make a copy**. Name your copy and save it to your Google Drive.
2.  **Open the Apps Script Editor:**
    *   In your copied Google Sheet, go to **Extensions > Apps Script**. This will open the script editor.
3.  **Review Configuration Sheets:**
    *   Open the "Instructions" tab in your copied Google Sheet for a detailed guide on setting up:
        *   **Your Question Bank Sheets** (e.g., `TF`, `MC`, `ES`)
        *   `SelectQuestions_Config`
        *   `CanvasQuiz_Config`
4.  **Initial Run & Authorization:**
    *   The first time you run any function from the "Extra Menu" (e.g., "Select Questions"), Google will ask for authorization.
    *   Review the permissions carefully and grant them.
5.  **Canvas API Token:**
    *   The first time you run "Create Quiz on Canvas", if a `CANVAS_API_TOKEN` is not found, you will be prompted to enter it.
    *   This token will be stored in Script Properties for future use.

## 🛠️ Usage

Once set up, use the "Extra Menu" in your Google Sheet:

1.  **Select Questions:**
    *   Configure the `SelectQuestions_Config` sheet.
    *   Run **Extra Menu > Select Questions**.
    *   The `Selected Questions` sheet will be populated.
2.  **Create Quiz on Canvas:**
    *   Ensure the `Selected Questions` sheet (or the sheet specified in `CanvasQuiz_Config`) has the questions you want.
    *   Configure the `CanvasQuiz_Config` sheet with your Canvas URL, Course ID, quiz title, and points.
    *   Run **Extra Menu > Create Quiz on Canvas**.
    *   A summary will appear. The quiz will be created as **UNPUBLISHED** in your Canvas course.

## 📂 Code Structure

The Apps Script project is organized into the following main files/modules:

*   **`CreateCanvasQuiz.gs`:** Contains the main `onOpen` function, global constants, and the `QuestionShuffler`, and `CanvasQuizCreator` modules/objects.
    *   `QuestionShuffler`: Handles the logic for selecting questions.
    *   `CanvasQuizCreator`: Handles the logic for creating quizzes on Canvas.
*   **`UiUtils.gs`:** Contains helper functions for displaying toast messages and dialogs (if you implemented this).

## 📊 Configuration Sheets Detailed

### `SelectQuestions_Config`

| Header                       | Description                                                                 |
| :--------------------------- | :-------------------------------------------------------------------------- |
| Sheet Name                   | Exact name of your question bank sheet tab.                                 |
| Include                      | Checkbox (TRUE/FALSE) to include this sheet.                                |
| Total Questions from Sheet   | Max questions to select from this sheet.                                    |
| Questions per Lecture        | Max questions per lecture from this sheet (0 = no per-lecture limit).       |
| Lecture Range Min            | Min lecture number/ID to consider.                                          |
| Lecture Range Max            | Max lecture number/ID to consider.                                          |
| Do Not Pick Again            | Checkbox (TRUE/FALSE) to mark selected questions as "do not pick" in source. |

### `CanvasQuiz_Config`

| Parameter                 | Description                                                                    |
| :------------------------ | :----------------------------------------------------------------------------- |
| `CANVAS_API_URL`          | Your Canvas instance URL (NO `/api/v1` at the end).                            |
| `CANVAS_API_TOKEN`        | (Optional) Your Canvas API Access Token.                                       |
| `COURSE_ID`               | Numerical ID of your Canvas course.                                            |
| `SHEET_NAME`              | Name of the sheet holding questions for Canvas (e.g., "Selected Questions").   |
| `QUIZ_TITLE`              | Title of the quiz in Canvas.                                                   |
| `POINTS_PER_QUESTION_TF`  | Points for True/False questions.                                               |
| `POINTS_PER_QUESTION_MC`  | Points for Multiple Choice/Answer questions.                                   |
| `POINTS_PER_QUESTION_ES`  | Points for Essay questions.                                                    |
| `POINTS_PER_QUESTION_DEFAULT` | Default points if ID prefix doesn't match TF, MC, or ES.                   |

### Question Bank Sheet Structure (e.g., "TF", "MC")

**Row 1 must be headers.**

| Header           | Description                                                                 |
| :--------------- | :-------------------------------------------------------------------------- |
| `ID`             | Unique ID (e.g., `TF001`, `MC_Topic_01`). Prefix determines Canvas type.    |
| `Do not pick`    | Checkbox (TRUE/FALSE).                                                      |
| `Lecture`        | Lecture number/ID.                                                          |
| `Question Text`  | Full question text.                                                         |
| `Correct Answer` | For TF: `TRUE`/`FALSE`. For MC: `A` or `A,C,D`. For ES: Optional model answer. |
| `Option A`       | Text for option A.                                                          |
| `Option B`       | Text for option B.                                                          |
| ...              | ... (Option C-F as needed)                                                  |
| *(Other cols)*   | For your reference (e.g., Topic, Tags).                                     |

## 📜 License

This project is licensed under the [MIT License](LICENSE) - see the `LICENSE` file for details.

## 👤 Author

[Jean-Louis Bru, Ph.D.](https://www.jlouisbru.com/)

Instructional Assistant Professor at [Chapman University](https://www.chapman.edu/)

## 🤝 Contributing

Contributions are welcome! [Thank you so much for your support!](https://ko-fi.com/louisfr)
