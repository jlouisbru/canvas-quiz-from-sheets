// ========= SPREADSHEET UI MENU =======================================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Quiz Tools')
    .addItem('Select Questions', 'runShuffleQuestions')
    .addSeparator()
    .addItem('Create Quiz on Canvas', 'runQuizCreation')
    .addToUi();
}
