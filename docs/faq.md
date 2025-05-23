# Frequently Asked Questions - Canvas Quiz Creator for Google Sheets

## 🌟 General Questions

### What is Canvas Quiz Creator for Google Sheets?
Canvas Quiz Creator is a Google Sheets tool that connects with Canvas LMS. It helps educators manage question banks within Google Sheets and then use those questions to automatically create quizzes directly in their Canvas courses.

### Is this tool free to use?
Yes, Canvas Quiz Creator for Google Sheets is completely free and open-source under the MIT License.

## 🚀 Setup Questions

### Do I need to know programming to use this tool?
No! It's designed for educators with no programming experience. Simply make a copy of our [Google Sheet Template](<LINK_TO_YOUR_SHARED_GOOGLE_SHEET_TEMPLATE_VIEW_ONLY>) *(<-- Replace this link!)* and follow the "Instructions" tab within it to get started.

### How do I find my Canvas Course ID?
Your Course ID is a number that appears in the URL when you are viewing your course page in Canvas. For example, in `https://yourschool.instructure.com/courses/12345`, the Course ID is `12345`.

### How do I get a Canvas API Token?
1.  In Canvas, go to **Account** (your profile) > **Settings**.
2.  Scroll down to **Approved Integrations** (you might need to enable this feature if it's your first time or if your institution manages it).
3.  Click **"+ New Access Token"**.
4.  Give it a "Purpose" (e.g., "Google Sheets Quiz Tool") and an optional "Expires" date (leaving it blank means it won't expire, but setting an expiration is good practice).
5.  Click **"Generate Token"**.
6.  **Copy the generated token immediately.** You will not be able to see it again after navigating away from that page.
The script will prompt you for this token if it's not found in the `CanvasQuiz_Config` sheet or your Script Properties. It will then store it securely in your Script Properties.

### What permissions does the script need when I first run it?
When you first run a function from the "Extra Menu," Google will ask for authorization. The script needs permission to:
*   **View and manage your spreadsheets in Google Drive:** To read your configuration sheets (`SelectQuestions_Config`, `CanvasQuiz_Config`), your question bank sheets, and to write to/clear the `Selected Questions` sheet.
*   **Connect to an external service:** This is required to communicate with the Canvas API to create quizzes in your Canvas course.
*   **Display and run third-party web content in prompts and sidebars:** Used for displaying `toast` messages (brief notifications) and dialogs (like the summary pop-ups).
The code is open-source on GitHub, so you can review what it does. By making a copy of the Google Sheet, the script becomes *your* copy and runs under your authority.

## 🛠️ Usage Questions

### How do I select which questions go into a Canvas quiz?
The process involves two main steps using the "Extra Menu":
1.  **Configure `SelectQuestions_Config`:** Specify which of your "Question Bank Sheets" (e.g., "TF", "MC") to use, how many questions to pull from each, filter by lecture range, etc.
2.  Run **Extra Menu > Select Questions**. This action reads your configurations and populates the `Selected Questions` sheet with a random selection of questions.
3.  **Configure `CanvasQuiz_Config`:** Set your Canvas URL, Course ID, desired Quiz Title, and points per question type. Ensure the `SHEET_NAME` parameter points to `Selected Questions` (or whichever sheet holds the questions you want to upload).
4.  Run **Extra Menu > Create Quiz on Canvas**. This takes the questions from the specified sheet and builds the quiz in Canvas.

### Can I use this tool for multiple Canvas courses?
Yes. The most straightforward way is to have a separate copy of this Google Spreadsheet for each Canvas course. This keeps the `CanvasQuiz_Config` (especially the `COURSE_ID`) distinct for each course.

### Are quizzes published to students immediately after creation?
No. The script creates all quizzes in Canvas as **UNPUBLISHED**. This is a safety measure, allowing you to go into your Canvas course, review the quiz settings, question content, and points, and then manually publish it when you are ready for students to access it.

### What types of questions can I create in Canvas with this tool?
The script currently supports creating the following Canvas question types based on the `ID` prefix in your question sheets:
*   `TF...`: Becomes a **True/False** question in Canvas.
*   `MC...`: Becomes a **Multiple Choice** question (or **Multiple Answers** if your `Correct Answer` column for that MC question contains comma-separated letters like `A,C`).
*   `ES...`: Becomes an **Essay Question** in Canvas.
*   Other prefixes: Will default to Multiple Choice/Multiple Answers based on `Correct Answer` format, using `POINTS_PER_QUESTION_DEFAULT`.

### How are points assigned to questions in the Canvas quiz?
Points are assigned based on your settings in the `CanvasQuiz_Config` sheet and the `ID` prefix of each question:
*   `POINTS_PER_QUESTION_TF` for questions with IDs starting with `TF`.
*   `POINTS_PER_QUESTION_MC` for questions with IDs starting with `MC`.
*   `POINTS_PER_QUESTION_ES` for questions with IDs starting with `ES`.
*   `POINTS_PER_QUESTION_DEFAULT` for any other questions.

### What if my True/False questions are showing up as "Multiple Choice" in Canvas or the wrong answer is selected?
This script is designed to create actual "True/False" type questions in Canvas.
*   Ensure your question IDs start with `TF`.
*   Ensure the `Correct Answer` column in your Google Sheet for these TF questions contains the word `TRUE` or `FALSE` (boolean values from checkboxes are also handled correctly).
If issues persist, check the script execution logs for messages about how the `correctAnswerFromSheet` value was processed for those TF questions.

### My Multiple Choice questions are being skipped with an error like "No valid answer options found." Why?
This error means the script could not find any text in your `Option A`, `Option B`, through `Option F` columns for those specific MC question rows in the sheet being used to create the Canvas quiz (usually `Selected Questions`).
*   **Verify:** Do these MC question rows have their answer choice text filled into these option columns (e.g., `Option A` has "Paris", `Option B` has "London")?
*   **Column Headers:** Are the column headers in your question sheet *exactly* `Option A`, `Option B`, etc., matching the constants defined in the script?

## 🔒 Technical & Security Questions

### How secure is my Canvas API Token?
Your Canvas API Token is stored securely in **Script Properties**. This storage is specific to your copy of the spreadsheet and your Google account. It is not visible to others if you share the spreadsheet file with "View" access. Only users with "Edit" access to your spreadsheet *and* the ability to open the Apps Script editor could potentially access Script Properties (this is standard Apps Script behavior).

### What data does this tool access from my Google Sheets?
The script reads data from:
*   Your configuration sheets (`SelectQuestions_Config`, `CanvasQuiz_Config`).
*   Your question bank sheets (as specified in `SelectQuestions_Config`).
*   The `Selected Questions` sheet (or as specified in `CanvasQuiz_Config` for uploads).
It writes data to:
*   The `Selected Questions` sheet (clearing and populating it).
*   Potentially your source question bank sheets if you enable the "Do Not Pick Again" feature (it checks the "Do not pick" box).

### What data does this tool send to Canvas?
To create a quiz, the script sends:
*   Quiz settings like title and total points (though total points are often recalculated by Canvas based on added questions).
*   For each question: its text, type, points, and for MC/TF questions, the answer options and which one(s) are correct.
The script **does not** pull any student data, grades, or other course content *from* Canvas with its current features.

## ⚖️ Privacy & Compliance Questions (General Guidance)

*(Disclaimer: This information is for guidance only and not legal advice. Always consult your institution's policies regarding data handling and third-party tools.)*

### Is this tool FERPA (or similar privacy regulation) compliant?
This tool acts as a bridge between two systems you already use and are authorized to access: Google Workspace (for Google Sheets) and Canvas LMS.
*   It transfers data (your question content) from your Google Sheet to your Canvas course.
*   It does not store your question data or Canvas data on any third-party servers outside of Google and Canvas.
*   It uses secure API connections (HTTPS) for communication with Canvas.

You, as the educator, are responsible for ensuring that your use of any tool, including this one, complies with your institution's specific data privacy policies, FERPA, GDPR, or other applicable regulations. This includes how you manage and share the Google Sheet containing your question data.

### Who can see the question data in my Google Sheet?
Anyone you explicitly share your Google Sheet with will be able to see its content, including your question banks and configurations. Ensure you only share your spreadsheet with authorized individuals who have a legitimate educational reason to access this information, adhering to your institution's data handling policies.

### Do I need special permission from my institution to use this tool?
While this tool leverages standard functionalities of Google Workspace and the Canvas API (which your institution already provides), it's always a good practice to be aware of your institution's policies regarding:
*   Use of third-party or custom scripts that interact with the LMS.
*   Generating Canvas API tokens and their approved uses.
Some institutions may have review processes for tools that access LMS data, even if the access is initiated by an authorized faculty member. When in doubt, check with your IT department or instructional technology support team.

---

**Remember to replace `<LINK_TO_YOUR_SHARED_GOOGLE_SHEET_TEMPLATE_VIEW_ONLY>` with the actual link to your Google Sheet template!**

This version should align well with the style you showed and provide clear, concise answers for users of your "Canvas Quiz Creator for Google Sheets."
