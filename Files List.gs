/**
 * Transfers ownership of files to "jamal.riley@mathnasium.com".
 * Function cannot be run by the "jamal.riley@mathnasium.com" account.
 * A time-based trigger must be run by each user who uploads files into the "SAM Files" folder.
 */
function transferFileOwnership() {
  const folderId = "1choRG8qg-ojmcwoM5FTie1gD4FdOA5eX"; // Folder ID for the "SAM Files" folder. The ID is at the end of the folder URL: "https://drive.google.com/drive/folders/FOLDER_ID"
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFiles();

  while (files.hasNext()) {
    const file = files.next();
    const fileOwnerEmail = file.getOwner().getEmail();
    if (fileOwnerEmail === USER && fileOwnerEmail !== "jamal.riley@mathnasium.com") file.setOwner("jamal.riley@mathnasium.com"); // If the user owns the file and it isn't Jamal, then set it to Jamal.
  }
}

/**
 * Allows user to click "Initialize File Ownership Transfer Protocol" menu option in the Clover Z Custom Menu.
 * Upon being clicked, it will create a trigger (under the non-Jamal user's account) to check for any files owned by the user.
 * If there are any files owned by the user, it will transfer ownership of files to "jamal.riley@mathnasium.com".
 */
function setFOTPTrigger() {
  // Alerts user before proceeding
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert("⚠️ You're about to transfer ownership of files", `As a heads up, this will automatically transfer ownership of current and future files owned by you (${USER}) to jamal.riley@mathnasium.com.\nThis will only transfer ownership of files in the "SAM Files" Google Drive folder and any of its subfolders.\n\nIf you wish to proceed, please click "Ok".`, SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);

  if (response === ui.Button.OK) ScriptApp.newTrigger("fileOwnershipTransferProtocol").timeBased().everyMinutes(1).create();
  else SpreadsheetApp.getActive().toast("File ownership transfer protocol canceled.", "", 3);
}

/**
 * Executes the "convertExcelFiles", "generateAppointyFilesList", and "generateWhenIWorkFilesList" functions in order.
 */
async function generateAllFilesLists() {
  await convertExcelFiles();
  await generateAppointyFilesList();
  await generateWhenIWorkFilesList();
}

/**
 * Executes the "handleAppointyAuto" and "handleWhenIWorkAuto".
 */
async function handleDataAuto() {
  await handleAppointyAuto();
  await handleWhenIWorkAuto();
}