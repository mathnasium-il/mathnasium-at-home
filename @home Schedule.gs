// Function-dependent cells that are subject to change positions upon future updates.
const SCHEDULE_DROPDOWN_CELL = "F1";
const SCHEDULE_COLS = "F4:M";
const FIRST_COL_TO_HIDE = "L";
const UNHIDE_COLS = "L:M";
const WEEKEND_COLS = ["F", "G", "H", "I", "J", "K"];
const WEEKDAY_COLS = [...WEEKEND_COLS, "L", "M"];

/**
 * Hides columns K & L if the "@home Schedule" dropdown is set to "Weekend".
 * Unhides columns K & L if the "@home Schedule" dropdown is set to "Weekday".
 */
function hideColumn(sheet, dropdownCell) {
  const dropdownCellVal = sheet.getRange(dropdownCell).getValue();
  if (dropdownCellVal === "Weekend") sheet.hideColumns(ABC_ARRAY.indexOf(FIRST_COL_TO_HIDE) + 1, 2);
  else sheet.unhideColumn(sheet.getRange(UNHIDE_COLS));
}

/**
 * Executes the "hideColumn" function for the "@home Schedule" worksheets of both SAM spreadsheets.
 */
function hideColumns() {
  hideColumn(ADMIN_SCHEDULE_SHEET, SCHEDULE_DROPDOWN_CELL);
  hideColumn(INSTRUCTOR_SCHEDULE_SHEET, SCHEDULE_DROPDOWN_CELL);
}

/**
 * Sets the value of each student cell in the @home Schedule with a customized dynamic range of students.
 * Each range of students are composed based on who is and isn't yet scheduled, who is actively working in a given time slot, and who can teach which level of mathematics.
 */
function setSessionDropdowns() {
  const scheduleMode = ADMIN_SCHEDULE_SHEET.getRange(SCHEDULE_DROPDOWN_CELL).getValue();
  const [colCount, timeValRow, timeABCArray] = scheduleMode === "Weekend" ?
    [6, [[10, "00"], [10, 30], [11, "00"], [11, 30], [12, "00"], [12, 30]], WEEKEND_COLS] :
    [8, [[15, "00"], [15, 30], [16, "00"], [16, 30], [17, "00"], [17, 30], [18, "00"], [18, 30]], WEEKDAY_COLS];

  const scheduleRange = ADMIN_SCHEDULE_SHEET.getRange(4, 4, ADMIN_SCHEDULE_SHEET.getLastRow() - 3, colCount).getValues();
  const dropdownRow = SCHEDULE_HELPER_3_SHEET.getRange("1:1").getValues().flat(); // 1500Instructor1-1,	1500Instructor1-2, etc. We put the full row so that when indexed, it returns the proper letter of the alphabet.
  const [rangeStart, rangeEnd] = [5, 13];
  const instructorRowIds = ADMIN_SCHEDULE_SHEET.getRange("A4:A").getValues().flat();

  for (let row = 0; row < scheduleRange.length; row++) {
    for (let col = 0; col < scheduleRange[0].length; col++) {
      const instructorNum = instructorRowIds[row];
      const [hour, mins] = [timeValRow[col][0], timeValRow[col][1]];

      const instructorDropdown = `${hour}${mins}${instructorNum}`;
      const instructorDropdownCol = ABC_ARRAY[dropdownRow.indexOf(instructorDropdown)];
      const instructorDropdownRange = (instructorDropdownCol === undefined) ? "" : SCHEDULE_HELPER_3_SHEET.getRange(`${instructorDropdownCol}${rangeStart}:${instructorDropdownCol}${rangeEnd}`);

      const cellA1Value = timeABCArray[col] + (row + 4); // row + 4 because the the first row of the instructorRowIds range is 4
      const cellRange = ADMIN_SCHEDULE_SHEET.getRange(cellA1Value);
      const instructorRule = SpreadsheetApp.newDataValidation().requireValueInRange(instructorDropdownRange).build();
      cellRange.setDataValidation(instructorRule);

      // const count = row * scheduleRange[0].length + col + 1;
      // console.log(`${count}: ${cellA1Value}'s rule: ${instructorDropdownCol}${rangeStart}:${instructorDropdownCol}${rangeEnd}`);
    }
  }
}

/**
 * Clears the @home schedule of any scheduled students.
 */
function clearScheduledStudents() {
  // setActiveWorksheet(ADMIN_SCHEDULE_SHEET);
  ADMIN_SCHEDULE_SHEET.getRange(SCHEDULE_COLS).clearContent(); // .clearNote();
}

/**
 * Generates a note containing a list of math levels that each scheduled instructor on the "@home Schedule" worksheet can teach.
 */
function setInstructorNotes() {
  ADMIN_SCHEDULE_SHEET.getRange("E4:E").clearNote(); // Clears any notes before proceeding

  const lastRowRange = STAFF_SHEET.getRange("A2:A").getValues();
  const instructors = STAFF_SHEET.getRange(2, 1, getLastRow(lastRowRange), STAFF_SHEET.getLastColumn()).getValues();
  const instructorFullNames = [];
  const staffDetails = [];

  for (const instructor of instructors) {
    const [firstNeme, lastName, pronoun1, pronoun2] = [instructor[0], instructor[1], instructor[4], instructor[5]];
    const fullName = firstNeme !== "" && lastName !== "" && pronoun1 !== "" && pronoun2 !== "" ? `${firstNeme} ${lastName} (${pronoun1}/${pronoun2})` : "";
    instructorFullNames.push([fullName]);

    staffDetails.push({
      fullName,
      canTeachGF: instructor[7],
      canTeachGeometry: instructor[8],
      canTeachAAT: instructor[9],
      canTeachPrecalc: instructor[10],
      canTeachCalc1: instructor[11],
      canTeachCalc2: instructor[12],
      canTeachStats: instructor[13],
      canTeachTestPrep: instructor[14],
    })
  }

  const scheduledInstructors = ADMIN_SCHEDULE_SHEET.getRange("D4:D").getValues();
  scheduledInstructors.forEach((row, i) => row.push(i + 4));
  const filteredScheduledInstructors = scheduledInstructors.filter(row => row[0] !== "");

  filteredScheduledInstructors.forEach(row => {
    const staffInfo = staffDetails.find(instructor => instructor.fullName === row[0]);
    const geometry = staffInfo.canTeachGeometry ? "Geometry" : "";
    const aat = staffInfo.canTeachAAT ? "Adv. Algebra/Trig" : "";
    const precalc = staffInfo.canTeachPrecalc ? "Precalculus" : "";
    const calc1 = staffInfo.canTeachCalc1 ? "Calculus 1/AB" : "";
    const calc2 = staffInfo.canTeachCalc2 ? "Calculus 2/BC" : "";
    const stats = staffInfo.canTeachStats ? "Statistics" : "";
    const testPrep = staffInfo.canTeachTestPrep ? "SAT/ACT Prep" : "";

    const mathLevels = [geometry, aat, precalc, calc1, calc2, stats, testPrep].filter(subject => subject !== '');
    const note = mathLevels.join("\n");

    ADMIN_SCHEDULE_SHEET.getRange(row[1], 5).setNote(note);
  })
}