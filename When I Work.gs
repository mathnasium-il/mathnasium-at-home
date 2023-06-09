// Function-dependent cells that are subject to change positions upon future updates.
const WHEN_I_WORK_FILE_CELL = "B1";
const WHEN_I_WORK_AUTO_IMPORT_CHECKBOX = "H1";

/**
 * Looks at uploaded files in the "S.A.M. Files" folder.
 * Detects whether or not it's an Excel file by looking at its file type.
 * For approved Excel files, it copies an equivalent converted Google Sheet and moves it to the "Converted Excel Files" folder.
 */
function convertExcelFiles() {
  // Folder IDs for the "SAM Files" and "Converted Excel Files" folders, respectively
  const [sourceFolderId, destinationFolderId] = ["1choRG8qg-ojmcwoM5FTie1gD4FdOA5eX", "1D10YAsXdfRRjz8mezzOFBLC273B2OUze"];
  const sourceFolder = DriveApp.getFolderById(sourceFolderId);
  const sourceFiles = sourceFolder.getFiles();
  let convertedFileCount = 0;

  while (sourceFiles.hasNext()) {
    const file = sourceFiles.next();
    const sourceFileName = file.getName();
    const fileBlob = file.getBlob();

    const newFile = {
      title: sourceFileName + "_converted",
      parents: [{ id: destinationFolderId }]
    };

    if (!sourceFileName.includes("_converted") && file.getMimeType() === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") {
      Drive.Files.insert(newFile, fileBlob, { convert: true });
      file.setName(`${sourceFileNameF}_converted`);
      convertedFileCount++;
    } // else console.log("This file cannot be imported.");
  }

  console.log(`${convertedFileCount} ${convertedFileCount === 1 ? "file has " : "files have"} been converted.`);
}

/**
 * Looks at uploaded files in the "Converted Excel Files" folder.
 * Detects whether or not it's an approved When I Work Google Sheet by A) looking at its file type, B) verifying the first header cell, and C) verifying the number of worksheets.
 * For approved When I Work Excel files, it renames them to "Shift Schedule: M/d/yy h:mm a" where M is the month, d is the day, yy is the year, h is the hour, mm is the minute, and a is AM/PM.
 * For approved When I Work Excel files, it will also add the adjusted names to the "Files List" worksheet, making it available as a dropdown item in the "When I Work Intake" worksheet.
 */
function generateWhenIWorkFilesList() {
  const folderId = "1D10YAsXdfRRjz8mezzOFBLC273B2OUze"; // Folder ID for the "Converted Excel Files" folder
  const files = DriveApp.getFolderById(folderId).getFiles();
  const filesList = [];

  // If the speficied column in the "Files List" worksheet already has file names in it, clear it before adding a new batch.
  if (getLastRow(FILES_LIST_SHEET.getRange("C:C").getValues()) > 0) FILES_LIST_SHEET.getRange("C:C").clearContent();

  while (files.hasNext()) {
    const file = files.next();
    if (file.getMimeType() !== "application/vnd.google-apps.spreadsheet") continue;
    let [fileId, fileName] = [file.getId(), file.getName()];

    const ss = SpreadsheetApp.openById(fileId);
    const sheets = ss.getSheets();
    const firstHeaderCell = sheets[0].getRange(1, 1).getValue(); // "Employee ID"

    const fileCreatedDate = file.getDateCreated();
    const formattedFileCreatedDate = Utilities.formatDate(fileCreatedDate, TIMEZONE, "M/d/yy h:mm a");

    if (firstHeaderCell === "Employee ID" && sheets.length === 2) {
      if (!fileName.includes("Shift Schedule: ")) {
        fileName = `Shift Schedule: ${formattedFileCreatedDate}`;
        file.setName(fileName);
      }
      filesList.push([fileName]);
    } // else console.log(`${fileName} is not an approved "When I Work" Google Sheet.`);
  }

  FILES_LIST_SHEET.getRange(1, 3, filesList.length, 1).setValues(filesList);
  const filesCountUnits = (filesList.length === 1) ? "Google Sheet has" : "Google Sheets have";
  console.log(`${filesList.length} ${filesCountUnits} been successfully added to the ${FILES_LIST_SHEET.getName()} worksheet.`);
}

/**
 * Imports When I Work file into S.A.M. spreadsheet.
 * Requires file name to import.
 * Will select the more recently created file if 2 files with the same name exist.
 */
function importWhenIWorkFile(fileName) {
  const fileId = DriveApp.getFilesByName(fileName).next().getId();
  const ss = SpreadsheetApp.openById(fileId);
  const scheduleSheet = ss.getSheets()[1];
  const scheduleSheetData = scheduleSheet.getRange(1, 1, scheduleSheet.getLastRow(), scheduleSheet.getLastColumn()).getValues();

  for (let i = 1; i < scheduleSheetData.length; i++) {
    const row = scheduleSheetData[i];
    const [shiftStart, shiftEnd] = [Number.isNaN(row[6]) ? "" : row[6], Number.isNaN(row[7]) ? "" : row[7]];
    shiftStart.setHours(shiftStart.getHours() - 1); // Set hours to correct time since shift times import off by 1 hour.
    shiftEnd.setHours(shiftEnd.getHours() - 1); // Set hours to correct time since shift times import off by 1 hour.
    [row[6], row[7]] = [shiftStart, shiftEnd];
  }

  WHEN_I_WORK_INTAKE_SHEET.getRange(2, 1, WHEN_I_WORK_INTAKE_SHEET.getLastRow(), WHEN_I_WORK_INTAKE_SHEET.getLastColumn()).clearContent();
  WHEN_I_WORK_INTAKE_SHEET.getRange(2, 1, 1, scheduleSheetData[0].length).setFontColor("#ffffff").setBackground("#589540").setFontWeight("bold").setHorizontalAlignment("center");
  WHEN_I_WORK_INTAKE_SHEET.getRange(2, 1, scheduleSheetData.length, scheduleSheetData[0].length).setValues(scheduleSheetData);
  SpreadsheetApp.getActive().toast(`${scheduleSheetData.length} ${scheduleSheetData.length === 1 ? "shift has" : "shifts have"} been successfully uploaded into the "${WHEN_I_WORK_INTAKE_SHEET.getName()}" worksheet.`, `${SUCCESS_NOTIF}, ${USER_FNAME}! ${EMOJI}`, 5);
}

/**
 * Sets active worksheet to "When I Work Intake".
 * Uses the shiftImporter function to import When I Work shift data.
 * Prompts user on whether or not to clear sessions upon import.
 */
async function handleWhenIWorkManual() {
  const fileName = WHEN_I_WORK_INTAKE_SHEET.getRange(WHEN_I_WORK_FILE_CELL).getValue();

  setActiveWorksheet(WHEN_I_WORK_INTAKE_SHEET.getName()); // Moves user to spreadsheet before proceeding
  const ui = SpreadsheetApp.getUi();
  const response1 = ui.alert("⚠️ You're about to import a new file", `As a heads up, you're about to import When I Work shift data from "${fileName}". Do you wish to proceed?`, SpreadsheetApp.getUi().ButtonSet.YES_NO); // Alerts user before proceeding

  if (response1 == ui.Button.YES) {
    const occupiedSlotCount = ADMIN_SCHEDULE_SHEET.getRange("B2").getValue();

    if (occupiedSlotCount > 0) {
      const response2 = ui.alert("💡 You're schedule isn't quite empty", "The schedule still has students on it. Do you want to clear the schedule?", SpreadsheetApp.getUi().ButtonSet.YES_NO);

      if (response2 === ui.Button.YES) {
        clearScheduledStudents();
        SpreadsheetApp.getActive().toast("You have successfully cleared the schedule!", `${SUCCESS_NOTIF}, ${USER_FNAME}! ${EMOJI}`, 5);
      } else SpreadsheetApp.getActive().toast("The schedule was not cleared.", "", 5);
    }
    await importWhenIWorkFile(fileName);
    await setInstructorNotes();
  } else SpreadsheetApp.getActive().toast("The When I Work file import was canceled.", "", 5);
}

/**
 * Uses the importWhenIWorkFile function to automatically import the most recently uploaded When I Work shift data.
 */
async function handleWhenIWorkAuto() {
  const isAutoImportEnabled = APPOINTY_INTAKE_SHEET.getRange(APPOINTY_AUTO_IMPORT_CHECKBOX).getValue();
  if (isAutoImportEnabled) {
    const fileName = WHEN_I_WORK_INTAKE_SHEET.getRange(WHEN_I_WORK_FILE_CELL).getValue();
    const fileId = DriveApp.getFilesByName(fileName).next().getId();

    const sheetDate = SpreadsheetApp.openById(fileId).getSheets()[0].getRange("D1").getValue();
    const a3 = WHEN_I_WORK_INTAKE_SHEET.getRange("A3").getValue();
    const isSameData = Utilities.formatDate(sheetDate, TIMEZONE, "MMM d, YYYY") === Utilities.formatDate(a3, TIMEZONE, "MMM d, YYYY");

    if (!isSameData) {
      await importWhenIWorkFile(fileName);
      await setInstructorNotes();
    }
    // else console.log("No new data to import.");
  } else return;
}

/**
 * Generates a list of staff with pronouns for the dropdown in the "When I Work Intake" worksheet.
 */
function generateWhenIWorkDropdowns() {
  const lastRowRange = STAFF_SHEET.getRange("A2:A").getValues();
  const staffMembers = STAFF_SHEET.getRange(2, 1, getLastRow(lastRowRange), STAFF_SHEET.getLastColumn()).getValues();
  const staffMemberFullNames = [];

  for (const staffMember of staffMembers) {
    const [firstNeme, lastName, pronoun1, pronoun2] = [staffMember[0], staffMember[1], staffMember[4], staffMember[5]];
    const fullName = firstNeme !== "" && lastName !== "" && pronoun1 !== "" && pronoun2 !== "" ? `${firstNeme} ${lastName} (${pronoun1}/${pronoun2})` : "";
    staffMemberFullNames.push([fullName]);
  }

  const whenIWorkHeaders = WHEN_I_WORK_INTAKE_SHEET.getRange(2, 1, 1, WHEN_I_WORK_INTAKE_SHEET.getLastColumn()).getValues().flat();
  const nameCIN = whenIWorkHeaders.indexOf("Name");
  const range = WHEN_I_WORK_INTAKE_SHEET.getRange(`${ABC_ARRAY[nameCIN]}3:${ABC_ARRAY[nameCIN]}`);
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(staffMemberFullNames.sort()).build();
  range.setDataValidation(rule);
}