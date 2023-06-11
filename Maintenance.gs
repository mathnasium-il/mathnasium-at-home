/**
 * Deletes files older than 48 hours from the "SAM Files" folder.
 */
function deleteDatedFiles() {
  const samFolderId = "1choRG8qg-ojmcwoM5FTie1gD4FdOA5eX"; // Folder ID for the "SAM Files" folder. The ID is at the end of the folder URL: "https://drive.google.com/drive/folders/FOLDER_ID"
  const samFolder = DriveApp.getFolderById(samFolderId);
  const samFiles = samFolder.getFiles();
  let deletedFileCount = 0;

  while (samFiles.hasNext()) {
    const file = samFiles.next();
    const fileCreatedDate = file.getDateCreated();
    const fileId = file.getId();

    const deadline = new Date();
    deadline.setDate(deadline.getDate() - 3);

    if (fileCreatedDate <= deadline) {
      Drive.Files.remove(fileId);
      deletedFileCount++;
    }
  }

  const convertedExcelsFolderId = '1D10YAsXdfRRjz8mezzOFBLC273B2OUze'; // Folder ID for the "Converted Excel Files" folder. The ID is at the end of the folder URL: "https://drive.google.com/drive/folders/FOLDER_ID"
  const convertedExcelsFolder = DriveApp.getFolderById(convertedExcelsFolderId);
  const convertedExcelsFiles = convertedExcelsFolder.getFiles();

  while (convertedExcelsFiles.hasNext()) {
    const file = convertedExcelsFiles.next();
    const fileCreatedDate = file.getDateCreated();
    const fileId = file.getId();

    const deadline = new Date();
    deadline.setDate(deadline.getDate() - 2);

    if (fileCreatedDate <= deadline) {
      Drive.Files.remove(fileId);
      deletedFileCount++;
    }
  }

  console.log(`${deletedFileCount} ${deletedFileCount === 1 ? "file has" : "files have"} been deleted.`);
}