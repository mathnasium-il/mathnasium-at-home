// Function-dependent cells that are subject to change positions upon future updates.
const APPOINTY_FILE_CELL = "B1";
const APPOINTY_AUTO_IMPORT_CHECKBOX = "H1";

/**
 * Looks at uploaded files in the "SAM Files" folder.
 * Detects whether or not it's an approved Appointy CSV file by A) looking at its file type and B) verifying the first header cell.
 * For approved Appointy CSV files that are not yet renamed, it renames them to "Appointy CSV: M/d/yy h:mm a" where M is the month, d is the day, yy is the year, h is the hour, mm is the minute, and a is AM/PM.
 * It will also add the adjusted names to the "Files List" worksheet, making it available as a dropdown item in the "Appointy Data Intake" worksheet.
 */
function generateAppointyFilesList() {
  const folderId = "1choRG8qg-ojmcwoM5FTie1gD4FdOA5eX"; // Folder ID for the "S.A.M. Files" folder. The ID is at the end of the folder URL: "https://drive.google.com/drive/folders/FOLDER_ID"
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFiles();
  const filesList = [];

  // If the speficied column in the "Files List" worksheet already has file names in it, clear it before adding a new batch.
  if (getLastRow(FILES_LIST_SHEET.getRange("A:A").getValues()) > 0) FILES_LIST_SHEET.getRange("A:A").clearContent();

  while (files.hasNext()) {
    const file = files.next();
    if (file.getMimeType() !== "text/csv") continue;
    let fileName = file.getName();
    const firstHeaderCell = Utilities.parseCsv(file.getBlob().getDataAsString())[0][0]; // "locationName"
    const fileCreatedDate = file.getDateCreated();
    const formattedFileCreatedDate = Utilities.formatDate(fileCreatedDate, TIMEZONE, "M/d/yy h:mm a");

    if (firstHeaderCell === "locationName") {
      if (!fileName.includes("Appointy: ")) {
        fileName = `Appointy: ${formattedFileCreatedDate}`;
        file.setName(fileName);
      }
      filesList.push([fileName]);
    } // else console.log(`${fileName} is not an approved "Appointy" CSV file.`);
  }

  FILES_LIST_SHEET.getRange(1, 1, filesList.length, 1).setValues(filesList);
  const filesCountUnits = (filesList.length === 1) ? "CSV file has" : "CSV files have";
  console.log(`${filesList.length} ${filesCountUnits} been successfully added to the ${FILES_LIST_SHEET.getName()} worksheet.`);
}

/**
 * Imports Appointy CSV file into SAM spreadsheet.
 * Requires file name to import.
 * Will select the more recently created file if 2 files with the same name exist.
 */
function importAppointyFile(fileName) {
  const file = DriveApp.getFilesByName(fileName).next();
  const dataBlob = file.getBlob();
  const csvString = dataBlob.getDataAsString();
  const csvData = Utilities.parseCsv(csvString);
  const csvDataUnits = (csvData.length === 1) ? "session has" : "sessions have";

  const b3 = APPOINTY_INTAKE_SHEET.getRange("B3").getValue();
  const isSameData = csvData[1][1].substring(0, 12) === Utilities.formatDate(b3, TIMEZONE, "MMM dd, YYYY");

  if (isSameData) {
    SpreadsheetApp.getActive().toast("No new data to import.", "", 5);
  } else {
    // Pastes CSV data into S.A.M. Spreadsheet
    APPOINTY_INTAKE_SHEET.getRange(2, 1, APPOINTY_INTAKE_SHEET.getLastRow(), APPOINTY_INTAKE_SHEET.getLastColumn()).clearContent(); // Clears current data with the exception of the first 2 rows
    APPOINTY_INTAKE_SHEET.getRange(2, 1, 1, csvData[0].length)
      .setFontColor("#ffffff")
      .setBackground("#ef3e33")
      .setFontWeight("bold")
      .setHorizontalAlignment("center");

    APPOINTY_INTAKE_SHEET.getRange(2, 1, csvData.length, csvData[0].length).setValues(csvData);

    SpreadsheetApp.getActive().toast(`${csvData.length} ${csvDataUnits} been successfully uploaded into the "${APPOINTY_INTAKE_SHEET.getName()}" worksheet.`, `${SUCCESS_NOTIF}, ${USER_FNAME}! ${EMOJI}`, 5);
  }
}

/**
 * Sets active worksheet to "Appointy Intake".
 * Uses the importAppointyFile function to import Appointy session data.
 * Prompts user on whether or not to clear sessions upon import.
 */
function handleAppointyManual() {
  const fileName = APPOINTY_INTAKE_SHEET.getRange(APPOINTY_FILE_CELL).getValue(); // Gets the name of the CSV file

  setActiveWorksheet(APPOINTY_INTAKE_SHEET.getName()); // Moves user to spreadsheet before proceeding
  const ui = SpreadsheetApp.getUi();
  const response1 = ui.alert("⚠️ You're about to import a new file", `As a heads up, you're about to import Appointy session data from "${fileName}". Do you wish to proceed?`, SpreadsheetApp.getUi().ButtonSet.YES_NO); // Alerts user before proceeding

  if (response1 === ui.Button.YES) {
    const occupiedSlotCount = ADMIN_SCHEDULE_SHEET.getRange("B2").getValue();

    if (occupiedSlotCount > 0) {
      const response2 = ui.alert("💡 You're schedule isn't quite empty", "The schedule still has students on it. Do you want to clear the schedule?", SpreadsheetApp.getUi().ButtonSet.YES_NO);

      if (response2 === ui.Button.YES) {
        clearScheduledStudents();
        SpreadsheetApp.getActive().toast("You have successfully cleared the schedule!", `${SUCCESS_NOTIF}, ${USER_FNAME}! ${EMOJI}`, 5);
      } else SpreadsheetApp.getActive().toast("The schedule was not cleared.", "", 5);
    }
    importAppointyFile(fileName);
  } else SpreadsheetApp.getActive().toast("The Appointy file import was canceled.", "", 3);
}

/**
 * Uses the importAppointyFile function to automatically import the most recently uploaded Appointy session data.
 */
function handleAppointyAuto() {
  const isAutoImportEnabled = APPOINTY_INTAKE_SHEET.getRange(APPOINTY_AUTO_IMPORT_CHECKBOX).getValue();
  if (isAutoImportEnabled) {
    const fileName = APPOINTY_INTAKE_SHEET.getRange(APPOINTY_FILE_CELL).getValue(); // Gets the name of the CSV file
    const file = DriveApp.getFilesByName(fileName).next();
    const dataBlob = file.getBlob();
    const csvString = dataBlob.getDataAsString();
    const csvData = Utilities.parseCsv(csvString);

    const csvDate = csvData[1][1].substring(0, 12);
    const b3 = APPOINTY_INTAKE_SHEET.getRange("B3").getValue();
    const isSameData = csvDate === Utilities.formatDate(b3, TIMEZONE, "MMM dd, YYYY");

    if (!isSameData) {
      APPTDATE_DROPDROWN.getRange("A1").setValue(`=VALUE("${Utilities.formatDate(new Date(csvDate), TIMEZONE, "M/d/YYYY")}")`);
      APPOINTY_INTAKE_SHEET.getRange(2, 1, APPOINTY_INTAKE_SHEET.getLastRow(), APPOINTY_INTAKE_SHEET.getLastColumn()).clearContent();
      APPOINTY_INTAKE_SHEET.getRange(2, 1, 1, csvData[0].length).setFontColor("#ffffff").setBackground("#ef3e33").setFontWeight("bold").setHorizontalAlignment("center");
      APPOINTY_INTAKE_SHEET.getRange(2, 1, csvData.length, csvData[0].length).setValues(csvData);

      const occupiedSlotCount = ADMIN_SCHEDULE_SHEET.getRange("B2").getValue();
      if (occupiedSlotCount > 0) clearScheduledStudents();
    } // else console.log("No new data to import.");
  } else return;
}

/**
 * Adds "Mathnasium@home" to the "locationName" column and "Appointment Confirmed" to the "status" column when students are manually added to "Appointy Intake" worksheet.
 */
function handleManualSession() {
  const lastRowRange = APPOINTY_INTAKE_SHEET.getRange("C3:C").getValues(); // "studentName" column
  const sessionData = APPOINTY_INTAKE_SHEET.getRange(3, 1, getLastRow(lastRowRange), APPOINTY_INTAKE_SHEET.getLastColumn()).getValues();
  sessionData.forEach((row, i) => row.push(i + 3));

  const APPOINTY_HEADERS = APPOINTY_INTAKE_SHEET.getRange(2, 1, 1, APPOINTY_INTAKE_SHEET.getLastColumn()).getValues().flat();
  const [LOCATION_NAME_CIN, APPT_DATE_CIN, STUDENT_NAME_CIN, SESSION_TYPE_CIN, STATUS_CIN] = [APPOINTY_HEADERS.indexOf("locationName"), APPOINTY_HEADERS.indexOf("appointmentDate"), APPOINTY_HEADERS.indexOf("studentName"), APPOINTY_HEADERS.indexOf("Session Type"), APPOINTY_HEADERS.indexOf("status")];

  const filteredSessionData = sessionData.filter(row => (row[LOCATION_NAME_CIN] === "" || row[STATUS_CIN] === "") && row[APPT_DATE_CIN] !== "" && row[STUDENT_NAME_CIN] !== "" && row[SESSION_TYPE_CIN] !== "");

  if (filteredSessionData.length > 0) {
    filteredSessionData.forEach(row => {
      APPOINTY_INTAKE_SHEET.getRange(row[row.length - 1], LOCATION_NAME_CIN + 1).setValue("Mathnasium@home");
      APPOINTY_INTAKE_SHEET.getRange(row[row.length - 1], STATUS_CIN + 1).setValue("Appointment Confirmed");
    });
  }
}

/**
 * Generates dropdown options for student names and session types in the "Appointy Intake" worksheet.
 */
function generateAppointyDropdowns() {
  const studentHeaders = STUDENTS_SHEET.getRange(1, 1, 1, STUDENTS_SHEET.getLastColumn()).getValues().flat();
  const [studentFirstCIN, studentLastCIN] = [studentHeaders.indexOf("First Name"), studentHeaders.indexOf("Last Name")];
  const studentHeadersArr = [studentFirstCIN, studentLastCIN]; // This array calculcates the left-most and right-most columns.
  const [minCol, maxCol] = [Math.min(...studentHeadersArr), Math.max(...studentHeadersArr)];

  const studentNameRange = STUDENTS_SHEET.getRange(`${ABC_ARRAY[minCol]}2:${ABC_ARRAY[maxCol]}`);
  const studentNameCols = studentNameRange.getValues();
  const [studentList, sessionTypeFullArr] = [[], []];
  const sessionTypeArr = ["K-8th", "Algebra 1", "Geometry", "Adv. Algebra/Trig", "Precalculus", "Calculus 1/AB", "Calculus 2/BC", "SAT/ACT Prep", "Statistics", "K-8th Private", "Algebra 1 Private", "Geometry Private", "Adv. Algebra/Trig Private", "Precalculus Private", "Calculus 1/AB Private", "Calculus 2/BC Private", "SAT/ACT Prep Private", "Statistics Private"];

  for (const student of studentNameCols) {
    const [firstNeme, lastName] = [student[0], student[1]];
    const fullName = firstNeme !== "" && lastName !== "" ? `${firstNeme} ${lastName}` : "";
    studentList.push([fullName]);
  }

  for (const sessionType of sessionTypeArr) {
    let [sessionType30, sessionType55] = [`${sessionType} - 30m`, `${sessionType} - 55m`];
    sessionTypeFullArr.push(sessionType30, sessionType55);
  }

  const appointyHeaders = APPOINTY_INTAKE_SHEET.getRange(2, 1, 1, APPOINTY_INTAKE_SHEET.getLastColumn()).getValues().flat();
  const [studentNameCIN, sessionTypeCIN] = [appointyHeaders.indexOf("studentName"), appointyHeaders.indexOf("Session Type")];

  const studentNameCellRange = APPOINTY_INTAKE_SHEET.getRange(`${ABC_ARRAY[studentNameCIN]}3:${ABC_ARRAY[studentNameCIN]}`);
  const sessionTypeCellRange = APPOINTY_INTAKE_SHEET.getRange(`${ABC_ARRAY[sessionTypeCIN]}3:${ABC_ARRAY[sessionTypeCIN]}`);

  const studentNameRule = SpreadsheetApp.newDataValidation().requireValueInList(studentList.sort()).build();
  const sessionTypeRule = SpreadsheetApp.newDataValidation().requireValueInList(sessionTypeFullArr).build();

  studentNameCellRange.setDataValidation(studentNameRule);
  sessionTypeCellRange.setDataValidation(sessionTypeRule);
}