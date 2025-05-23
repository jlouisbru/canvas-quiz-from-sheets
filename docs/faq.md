# Frequently Asked Questions (FAQ) - Canvas Quiz Creator for Google Sheets

## getSheetByNameeneral & Setup

**Q1: How do I get started with this tool?**
A: Please refer to the main [README.md](https://github.com/jlouisbru/canvas-quiz-from-sheets/blob/main/README.md) for detailed setup instructions. The key steps are:
    1. Make a copy of the [Google Sheet Template](<LINK_TO_YOUR_SHARED_GOOGLE_SHEET_TEMPLATE_VIEW_ONLY>). *(Remember to replace this link!)*
    2. Open **Extensions > Apps Script** in your copied sheet.
    3. Configure the `SelectQuestions_Config` and `CanvasQuiz_Config` sheets as per the "Instructions" tab in the Google Sheet.
    4. Run the functions from the "Extra Menu". You'll be asked for script authorization the first time.

**Q2: The script is asking for a lot of permissions. Is it safe?**
A: Yes, the script requires these permissions to function:
    *   **View and manage your spreadsheets in Google Drive:** Needed to read your configuration sheets (e.g., `SelectQuestions_Config`, `CanvasQuiz_Config`), your question bank sheets, and to write to the `Selected Questions` sheet.
    *   **Connect to an external service:** Needed to communicate with the Canvas API to create quizzes.
    *   **Display and run third-party web content in prompts and sidebars:** Used for displaying `toast` messages and potentially other UI elements.
    The code is open source on GitHub, so you can review what it does. When you make a copy of the Google Sheet, the script becomes *your* copy, running under your authority.

**Q3: Where do I get a Canvas API Token?**
A: In your Canvas instance:
    1. Go to **Account** (usually your profile picture in the sidebar).
    2. Click on **Settings**.
    3. Scroll down to **Approved Integrations**.
    4. Click the **"+ New Access Token"** button.
    5. Give it a purpose (e.g., "Google Sheets Quiz Creator") and an expiration date (optional, but recommended for security).
    6. Click **"Generate Token"**.
    7. **Important:** Copy the generated token immediately. You will NOT be able to see it again.
    The script will prompt you for this token if it's not found in the `CanvasQuiz_Config` sheet or your Script Properties. It will then store it securely in your Script Properties for future use.

**Q4: Do I need to know how to code to use this?**
A: No! The tool is designed to be used through the Google Sheets interface and its configuration sheets. You only need to interact with the Apps Script code if you want to customize it or contribute to its development.

## ❓ Question Selection (`SelectQuestions_Config`)

**Q5: Why aren't any questions being selected into the "Selected Questions" sheet?**
A: Check the following in your `SelectQuestions_Config` sheet and your question bank sheets:
    *   **Include Checkbox:** Is the "Include" box checked for the source sheets you want to use?
    *   **Sheet Name:** Does the "Sheet Name" in `SelectQuestions_Config` *exactly* match the actual tab name of your question bank sheet (case-sensitive)?
    *   **Lecture Range:** Are your `Lecture Range Min` and `Lecture Range Max` values set correctly for the lecture numbers/IDs in your question bank? If a question's lecture value falls outside this range, it won't be selected.
    *   **"Do not pick" Column:** In your question bank sheet, ensure the "Do not pick" column is not checked for questions you want to be eligible.
    *   **Sufficient Questions:** Do you have enough questions in your source sheets that meet all criteria (lecture range, not marked "Do not pick") to fulfill the "Total Questions from Sheet" and "Questions per Lecture" settings?
    *   **Logs:** Check **Extensions > Apps Script > Executions** for any specific error messages or warnings.

**Q6: What does "Do Not Pick Again" do?**
A: If you check the "Do Not Pick Again" box in `SelectQuestions_Config` for a particular source sheet, any questions selected from that sheet will then have their own "Do not pick" checkbox automatically checked in that original source sheet. This helps prevent you from selecting the same questions repeatedly for different quizzes. It uses the `ID` column to match questions.

## 캔버스 Canvas Quiz Creation (`CanvasQuiz_Config`)

**Q7: My quiz wasn't created on Canvas / I got an API error. What should I check?**
A: Common issues include:
    *   **`CANVAS_API_URL`:** Ensure this is your base Canvas URL (e.g., `https://yourschool.instructure.com`) and does **NOT** end with `/api/v1`.
    *   **`CANVAS_API_TOKEN`:** Is the token correct and still valid (not expired)? If you left it blank, ensure you entered it correctly when prompted.
    *   **`COURSE_ID`:** Is this the correct numerical ID for your Canvas course? You find it in the URL of your course page (e.g., `.../courses/12345`).
    *   **Canvas Permissions:** Does the user whose API token you're using have permission to create quizzes in that course?
    *   **Logs:** Check **Extensions > Apps Script > Executions** for detailed error messages from the Canvas API. These often provide clues.

**Q8: The points for my questions are wrong in Canvas, or a question type is incorrect.**
A:
    *   **Points:** Check your `CanvasQuiz_Config` sheet. Ensure `POINTS_PER_QUESTION_TF`, `_MC`, `_ES`, and `_DEFAULT` are set to the correct numeric values. The script uses the prefix of the `ID` column (e.g., "TF_", "MC_", "ES_") in your `Selected Questions` sheet to determine which point value to apply.
    *   **Question Type:**
        *   IDs starting with `TF` become "True/False" questions.
        *   IDs starting with `MC` become "Multiple Choice" (or "Multiple Answers" if the `Correct Answer` column has comma-separated letters like `A,C`).
        *   IDs starting with `ES` become "Essay" questions.
        *   If an `ID` doesn't have one of these prefixes, it will likely default to "Multiple Choice" or "Multiple Answers" based on the `Correct Answer` format.

**Q9: My True/False questions are all defaulting to "True" as the correct answer in Canvas.**
A: This usually happens if the `Correct Answer` column in your `Selected Questions` (or source) sheet for TF questions is not being interpreted correctly as boolean `true` or `false`, or the string "TRUE" or "FALSE".
    *   Ensure the `Correct Answer` column for TF questions in your sheet contains the word `TRUE` or `FALSE` (case-insensitive, boolean values from checkboxes are also fine).
    *   The script is designed to convert these to uppercase "TRUE" or "FALSE" and set the Canvas question accordingly. If issues persist, check the script logs for how it's processing the `correctAnswerFromSheet` value for those TF questions.

**Q10: My Multiple Choice questions are being skipped with "No valid answer options found."**
A: This means the script could not find text in the `Option A`, `Option B`, etc., columns for those MC question rows in your `Selected Questions` sheet.
    *   Verify that your MC questions have their answer choice text filled into these option columns.
    *   Ensure the column headers (`Option A`, `Option B`, etc.) in your `Selected Questions` sheet exactly match the constants defined in the script (`COL_OPTION_A`, `COL_OPTION_B`, etc.).

**Q11: Where is the quiz in Canvas after the script runs?**
A: The quiz is created as **UNPUBLISHED** in the Canvas course you specified by `COURSE_ID`. You need to go to your course in Canvas, find the quiz (usually under "Quizzes"), review it, and then publish it manually for students to see.

## 💻 Script & Code

**Q12: Can I modify the script code?**
A: Yes, it's Google Apps Script (JavaScript). You can access it via **Extensions > Apps Script**. However, be cautious when making changes unless you are familiar with JavaScript and Apps Script, as incorrect modifications can break its functionality. It's recommended to make a backup copy of your spreadsheet (and thus the bound script) before making significant code changes.

**Q13: How do I get updates to the script if you release a new version?**
A: Since you make a copy of the Google Sheet, your script is independent. To get updates:
    1. You would need to check the [GitHub repository](https://github.com/jlouisbru/canvas-quiz-from-sheets) for new versions.
    2. If there's a new version, you would typically need to manually copy the updated code from the new `.gs` files on GitHub into your Apps Script editor, replacing your existing code.
    3. Alternatively, if you use `clasp`, you could pull the latest changes from the GitHub repository to your local machine and then `clasp push` them to your script project.
    Review the `CHANGELOG.md` for details on what changed in new versions.

---

**Remember to replace `<LINK_TO_YOUR_SHARED_GOOGLE_SHEET_TEMPLATE_VIEW_ONLY>` with the actual link!**

You can add more questions as they come up or as you anticipate user needs. This should be a solid starting point.
