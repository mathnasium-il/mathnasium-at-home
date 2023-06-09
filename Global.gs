const SAM_ADMIN_SS = SpreadsheetApp.openById("1HyoDmMikyeNhej47qtbKUt5o1fGDv2E4R-hu3x0QNfk");
const FILES_LIST_SHEET = SAM_ADMIN_SS.getSheetByName("Files List");
const APPTDATE_DROPDROWN = SAM_ADMIN_SS.getSheetByName("appointmentDate Dropdown");
const SCHEDULE_HELPER_3_SHEET = SAM_ADMIN_SS.getSheetByName("@home Schedule Helper 3");
const STUDENTS_SHEET = SAM_ADMIN_SS.getSheetByName("Students");
const STAFF_SHEET = SAM_ADMIN_SS.getSheetByName("Staff");
const ADMIN_SCHEDULE_SHEET = SAM_ADMIN_SS.getSheetByName("@home Schedule");
const APPOINTY_INTAKE_SHEET = SAM_ADMIN_SS.getSheetByName("Appointy Intake");
const WHEN_I_WORK_INTAKE_SHEET = SAM_ADMIN_SS.getSheetByName("When I Work Intake");

const SAM_INSTRUCTOR_SS = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1liiDRP5AiOSHCFLFDHb5GF-ZGJTAqJKRmtLSDrCRPjc/edit#gid=642293000");
const INSTRUCTOR_SCHEDULE_SHEET = SAM_INSTRUCTOR_SS.getSheetByName("@home Schedule");

const TIMEZONE = "America/Chicago";
const D = new Date();

const ABC_ARRAY = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL", "BM", "BN", "BO", "BP", "BQ", "BR", "BS", "BT", "BU", "BV", "BW", "BX", "BY", "BZ", "CA", "CB", "CC", "CD", "CE", "CF", "CG", "CH", "CI", "CJ", "CK", "CL", "CM", "CN", "CO", "CP", "CQ", "CR", "CS", "CT", "CU", "CV", "CW", "CX", "CY", "CZ", "DA", "DB", "DC", "DD", "DE", "DF", "DG", "DH", "DI", "DJ", "DK", "DL", "DM", "DN", "DO", "DP", "DQ", "DR", "DS", "DT", "DU", "DV", "DW", "DX", "DY", "DZ", "EA", "EB", "EC", "ED", "EE", "EF", "EG", "EH", "EI", "EJ", "EK", "EL", "EM", "EN", "EO", "EP", "EQ", "ER", "ES", "ET", "EU", "EV", "EW", "EX", "EY", "EZ", "FA", "FB", "FC", "FD", "FE", "FF", "FG", "FH", "FI", "FJ", "FK", "FL", "FM", "FN", "FO", "FP", "FQ", "FR", "FS", "FT", "FU", "FV", "FW", "FX", "FY", "FZ"];

const SUCCESS_NOTIFS_LIST = ['"One"-derful', "Nice work", "Amazing", "Fantastic", "Impressive", "Not bad", "Math-tastic", "Woohoo", "Bravo"];
const NUM_OF_SUCCESS_NOTIFS = SUCCESS_NOTIFS_LIST.length;
const SUCCESS_NOTIF = SUCCESS_NOTIFS_LIST[Math.floor(Math.random() * NUM_OF_SUCCESS_NOTIFS)];

const EMOJIS_LIST = ["😀", "😃", "😄", "😁", "😊", "🤓", "😎", "✅"];
const NUM_OF_EMOJIS = EMOJIS_LIST.length;
const EMOJI = EMOJIS_LIST[Math.floor(Math.random() * NUM_OF_EMOJIS)];

const PERSONNEL = new Map();
PERSONNEL.set("jana.frank@mathnasium.com", "Jana");
PERSONNEL.set("anthony.paparo@mathnasium.com", "Anthony");
PERSONNEL.set("jamal.riley@mathnasium.com", "Jamal");
PERSONNEL.set("cloverz@mathnasium.com", "Laura");
PERSONNEL.set("oakparkriverforest@mathnasium.com", "Nancy");
PERSONNEL.set("lagrange@mathnasium.com", "Caitlyn");
PERSONNEL.set("mountprospect@mathnasium.com", "Brian");
const USER = Session.getActiveUser().getEmail();
const USER_FNAME = PERSONNEL.get(USER);

/**
 * 
 */
function menu() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Clover Z")
    .addItem("Clear schedule", "clearScheduledStudents")
    .addSubMenu(SpreadsheetApp.getUi().createMenu("Manually import data")
      .addItem("Import Appointy data", "handleAppointyManual")
      .addItem("Import When I Work data", "handleWhenIWorkManual"))
    .addSubMenu(SpreadsheetApp.getUi().createMenu("Manual override options")
      .addItem("Manually update session dropdowns", "setSessionDropdowns")
      .addItem("Manually update file dropdowns", "generateAllFilesLists"))
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu("About")
      .addItem("Documentation", "showDocumentation")
      .addItem("Initialize File Ownership Transfer Protocol", "setFOTPTrigger"))
    .addToUi();
}

/**
 * Helper function to get the last filled row of a given range. 
 */
function getLastRow(range) {
  const flatRange = range.flat();
  let flatRangeLength = flatRange.length;
  while (flatRange[flatRangeLength - 1] === "") {
    flatRange.splice(flatRangeLength - 1, 1);
    flatRangeLength--;
  }
  return flatRangeLength;
}

/**
 * Sets the active worksheet to the specified worksheet.
 */
const setActiveWorksheet = (sheetName) => SpreadsheetApp.getActive().getSheetByName(sheetName).activate();

/**
 * Shows a UI of SAM-related documentation for users to access.
 */
function showDocumentation() {
  const [name1, url1] = ["✏️ Admin Documentation", "https://docs.google.com/document/d/1endf0VCeT3Q0x3v9INlHGjc_ORMtbBPeZoes40sy7wE/edit?usp=sharing"];
  const html1 = `<html><body><p><a href="${url1}" target="blank" onclick="google.script.host.close()" style="font-family:arial; color:#ef3e33; text-decoration:none">${name1}</a></p></body></html>`;

  const [name2, url2] = ["✏️ Instructor Documentation", "https://docs.google.com/document/d/1endf0VCeT3Q0x3v9INlHGjc_ORMtbBPeZoes40sy7wE/edit?usp=sharing"];
  const html2 = `<html><body><p><a href="${url2}" target="blank" onclick="google.script.host.close()" style="font-family:arial; color:#000000; text-decoration:none">${name2}</a></p></body></html>`;

  const ui = HtmlService.createHtmlOutput(html1 + html2);
  SpreadsheetApp.getUi().showModelessDialog(ui, "Documentation");
}

/**
 * Executed certain onEdit functions based on the active worksheet.
 */
function updateOnEdit() {
  if (SpreadsheetApp.getActiveSheet().getName() === APPOINTY_INTAKE_SHEET.getSheetName()) handleManualSession();
  else if (SpreadsheetApp.getActiveSheet().getName() === ADMIN_SCHEDULE_SHEET.getSheetName()) hideColumns();
  else if (SpreadsheetApp.getActiveSheet().getName() === STUDENTS_SHEET.getSheetName()) generateAppointyDropdowns();
  else if (SpreadsheetApp.getActiveSheet().getName() === STAFF_SHEET.getSheetName()) generateWhenIWorkDropdowns();
}