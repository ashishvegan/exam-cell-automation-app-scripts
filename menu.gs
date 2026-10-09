function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("ESE")
    .addItem("Send Timetable", "sendExamDutyEmails")
    .addToUi();
}