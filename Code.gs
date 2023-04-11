const CATEGORY_COLUMN = 1;
const CATEGORY_FIRST_ROW = 2;

const TASK_TABLE_START_COLUMN = 3;
const TASK_TABLE_START_ROW = 3;
const TASK_TABLE_WIDTH = 4;
const TASK_CATEGORY_COLUMN = 3;
const TASK_TITLE_COLUMN = 4;
const TASK_STATUS_COLUMN = 5;
const TASK_READY_COLUMN = 6;

const CALENDAR_HEADER_DAY_ROW = 1;
const CALENDAR_HEADER_DATE_ROW = 2;
const CALENDAR_BODY_FIRST_ROW = 3;
const CALENDAR_BODY_ROW_TYPE_COL = 1;
const CALENDAR_BODY_TITLE_COL = 2;
const CALENDAR_BODY_NUM_DAYS_COL = 3;
const CALENDAR_BODY_FIRST_DATE_COL = 5;

const ss = SpreadsheetApp.getActiveSpreadsheet();
const masterList = ss.getSheetByName("Master list");
let thisMonthSheet = ss.getSheets()[1];

/**
 * @param {Date} todayDt - Date object for the date on which the sheet is being prepared
 */
function prepareNewSheet(todayDt) {
  const month = new Month(todayDt.getMonth()).name;
  const year = todayDt.getFullYear();

  // 1. Create a new sheet and name it [Month] [Year]
  const newSheetName = `${month} ${year}`;
  ss.insertSheet(newSheetName, 1);
  thisMonthSheet = ss.getSheetByName(newSheetName);

  // 2. Populate with all the ready, not-done tasks, like this:
  //    ###############################
  //    # A. WORK                     #
  //    ###############################
  //    Web app    0 (formula to add up ticks)
  //    Lula       "
  //    etc.
  const lastRow = masterList.getLastRow();

  if (lastRow < CATEGORY_FIRST_ROW) {
    Logger.log("There are no categories");
    return;
  }

  // [ ["Work", "Web app", false, true], ["Books", "Americanah", true, true], ... ]
  const tasks = masterList
    .getRange(TASK_TABLE_START_ROW,TASK_TABLE_START_COLUMN, lastRow - TASK_TABLE_START_ROW + 1, TASK_TABLE_WIDTH)
    .getValues()
    .filter(([category, title, completed, ready]) => !completed && ready);

  // We only care about categories that have incomplete tasks
  // e.g. [ "Work", "Books", ... ]
  let activeCategories = masterList
    .getRange(CATEGORY_FIRST_ROW, CATEGORY_COLUMN, lastRow - CATEGORY_FIRST_ROW + 1, 2)
    .getValues()
    .filter(([category, activeTaskCount]) => activeTaskCount > 0)
    .map(([category]) => category);

  let row = CALENDAR_BODY_FIRST_ROW;
  let col = CALENDAR_BODY_TITLE_COL;

  activeCategories.forEach((category, index) => {
    createHeaderRow(row, index, category);
    row++;
    
    tasks
      .filter(task => task[0] === category)
      .forEach(task => {
        // Put "task" into the first column
        thisMonthSheet
          .getRange(row, col - 1)
          .setValue("task");
        // Next column has the task's title
        thisMonthSheet
          .getRange(row, col)
          .setValue(task[1]);
        // Next column has the formula that calculates how many days were spent
        // working on the task
        thisMonthSheet
          .getRange(row, col + 1)
          .setHorizontalAlignment("center")
          // R[0]C[1] will be a blank column
          // R[0]C[2] will be the first date column
          .setFormulaR1C1("=COUNTIF(R[0]C[1]:R[0]C[2], TRUE)");
        row++;
      });

    row++; // blank line
  });

  // 3. Resize columns
  //  3.1. Resize task title column to fit longest task title
  thisMonthSheet.autoResizeColumn(CALENDAR_BODY_TITLE_COL);
  //  3.2. Next two columns should be 35 (number of days + blank columns)
  thisMonthSheet.setColumnWidths(CALENDAR_BODY_TITLE_COL + 1, 2, 35);

  // 4. Hide the previous month's sheet
  ss.getSheets()[2]?.hideSheet();

  // 5. Hide the row type column
  thisMonthSheet.hideColumn(
    thisMonthSheet.getRange(
      1,
      CALENDAR_BODY_ROW_TYPE_COL,
      thisMonthSheet.getMaxRows(),
      1,
    ),
  );
}

/**
 * @param {Date} dt - Date object for the date on which the column is being prepared
 */
function prepareNewColumn(dt) {
  const d = dt.getDate(); // e.g. 1, 2, 3, ...
  const ddd = new Day(dt.getDay()).abbreviation; // e.g. M, Tu, W, ...

  // 1. Add a new column to the left of the most recent date
  // (only if it is after the first day of the month because on Day 1,
  // the first date column is already included in the merged header row)
  const isFirstDay = d === 1;
  if (!isFirstDay) {
    thisMonthSheet.insertColumnsBefore(CALENDAR_BODY_FIRST_DATE_COL, 1);
  }

  // 2. Header rows:
  //    2.1. Day of the week (M, Tu, W, Th, etc)
  //    2.2. Date as d MMM
  thisMonthSheet
    .getRange(CALENDAR_HEADER_DAY_ROW, CALENDAR_BODY_FIRST_DATE_COL)
    .setFontSize(8)
    .setHorizontalAlignment("center")
    .setValue(ddd);

  thisMonthSheet
    .getRange(CALENDAR_HEADER_DATE_ROW, CALENDAR_BODY_FIRST_DATE_COL)
    .setValue(dt)
    .setNumberFormat("d MMM")
    .setHorizontalAlignment("center");

  // 3. For all incomplete tasks, add checkbox data validation
  // in the new column
  const lastRow = thisMonthSheet.getLastRow();
  for (let row = CALENDAR_BODY_FIRST_ROW; row <= lastRow; row++) {
    let rowType = thisMonthSheet
      .getRange(row, CALENDAR_BODY_ROW_TYPE_COL)
      .getValue();
    if (rowType === "task") {
      let titleCell = thisMonthSheet.getRange(row, CALENDAR_BODY_TITLE_COL);
      let checkboxCell = thisMonthSheet.getRange(row, CALENDAR_BODY_FIRST_DATE_COL);
      if (titleCell.getFontLine() === "line-through") {
        // This is a completed task
        checkboxCell.removeCheckboxes();
      } else {
        // not actually necessary; inserting column does this
        checkboxCell
          .insertCheckboxes() 
          .setHorizontalAlignment("center");
      }
    }
  }

  // 4. Date column should be 60
  thisMonthSheet.setColumnWidth(CALENDAR_BODY_FIRST_DATE_COL, 60);
}

/**
 * When a task is marked as done on the master list sheet, strike it out
 * and gray it out on the active calendar sheet.
 * @param e - the event object passed to the function by the edit trigger
 */
function onTaskDone(e) {
  const sheetName = e.range.getSheet().getName();
  if (sheetName !== "Master list") return;

  const row = e.range.getRow();
  const column = e.range.getColumn();
  if (column === TASK_STATUS_COLUMN) {
    const details = masterList
      .getRange(row, TASK_CATEGORY_COLUMN, 1, 2)
      .getValues()[0]; // e.g. ["Work", "Web app"]
    
    const completed = e.value === "TRUE" ? true : false;
    for (let i = CALENDAR_BODY_FIRST_ROW; i <= thisMonthSheet.getLastRow(); i++) {
      let row = thisMonthSheet.getRange(i, CALENDAR_BODY_TITLE_COL, 1, 2);
      if (row.getValue() === details[1]) {
        if (completed) {
          row
            .setFontLine("line-through")
            .setFontColor("lightgray");
        } else {
          row
            .setFontLine("none")
            .setFontColor("black");
        }

        break;
      }
    }
  }
}

/**
 * When a new task is added to the master list and marked as ready, create
 * a row for it in the active calendar sheet
 * @param e - the event object passed to the function by the edit trigger
 */
function onNewTaskAdded(e) {
  const sheetName = e.range.getSheet().getName();
  if (sheetName !== "Master list") return;

  const row = e.range.getRow();
  const column = e.range.getColumn();
  if (column === TASK_READY_COLUMN && e.value === "TRUE" ? true : false) {
    const details = masterList
      .getRange(row, TASK_CATEGORY_COLUMN, 1, 2)
      .getValues()[0];
    insertTask(details[0], details[1]);
  }
}

/**
 * @param {string} taskCategory - e.g. "Work", "Personal"
 * @param {string} taskTitle - The title of the task
 */
function insertTask(taskCategory, taskTitle) {
  // 1. Determine which categories are on the active calendar sheet
  // (Note that we cannot rely on the active task count for this because
  // maybe one category started out the month active, but during the month
  // its only active task was completed, rendering the category no longer
  // active).
  // activeCategories will look something like [[3, "Work"], [10, "Personal"], [15, "Books"], ...]
  const activeCategories = thisMonthSheet
    .getRange(
      CALENDAR_BODY_FIRST_ROW, CALENDAR_BODY_ROW_TYPE_COL,
      thisMonthSheet.getLastRow() - CALENDAR_BODY_FIRST_ROW + 1, 2)
    .getValues()
    // Capture the row number of each row
    .map((row, index) => [CALENDAR_BODY_FIRST_ROW + index, ...row])
    .filter(([rowNum, rowType]) => rowType === "category")
    // Remove the first 3 characters (e.g. "A. ", "B. ")
    .map(([rowNum, rowType, category]) => [rowNum, `${category.substring(3)}`]);

  // 2. Figure out where to insert the task on the calendar sheet
  // If the active calendar sheet does not contain the new task's category, we must
  // create a new section for the category
  const categoryIndex = activeCategories.findIndex(([rowNum, category]) => taskCategory === category);
  let mustCreateNewSection = categoryIndex < 0;
  let taskRow; // the row number where the new task will be inserted

  if (mustCreateNewSection) {
    // Create the section right at the end
    const lastRow = thisMonthSheet.getLastRow(); // last task recorded
    // + 1 --> blank line; + 2 --> new category; + 3 --> new task
    createHeaderRow(lastRow + 2, activeCategories.length, taskCategory);
    taskRow = lastRow + 3;
  } else {
    // Nice and easy: the calendar already contains this category
    if (categoryIndex === (activeCategories.length - 1)) { // if it's in the last category
      taskRow = thisMonthSheet.getLastRow() + 1;
    } else {
      // Get the category after taskCategory, and go up two rows in order to
      // get the row number of the last task inside taskCategory
      const lastTaskRow = activeCategories[categoryIndex + 1][0] - 2;
      thisMonthSheet.insertRowsAfter(lastTaskRow, 1);
      taskRow = lastTaskRow + 1;
    }
  }

  thisMonthSheet
    .getRange(taskRow, CALENDAR_BODY_ROW_TYPE_COL)
    .setValue("task");
  thisMonthSheet
    .getRange(taskRow, CALENDAR_BODY_TITLE_COL)
    .setValue(taskTitle);
  thisMonthSheet
    .getRange(taskRow, CALENDAR_BODY_NUM_DAYS_COL)
    .setFormulaR1C1("=COUNTIF(R[0]C[1]:R[0]C[2], TRUE)")
    .setHorizontalAlignment("center");
  thisMonthSheet
    .getRange(taskRow, CALENDAR_BODY_FIRST_DATE_COL)
    .insertCheckboxes()
    .setHorizontalAlignment("center");

  // 3. Remove checkboxes from previous days in the month (the task did not exist yet)
  const lastCol = thisMonthSheet.getLastColumn();
  if (lastCol > CALENDAR_BODY_FIRST_DATE_COL) {
    thisMonthSheet
      .getRange(taskRow, CALENDAR_BODY_FIRST_DATE_COL + 1, 1, lastCol - CALENDAR_BODY_FIRST_DATE_COL)
      .removeCheckboxes();
  }

  // 4. In case taskTitle is long...
  thisMonthSheet.autoResizeColumn(CALENDAR_BODY_TITLE_COL);
}

/**
 * @param {number} row - Which row the category heading is being added to
 * @param {number} index - The category's rank/position in the categories list (zero-indexed)
 * @param {string} category - The category's title
 */
function createHeaderRow(row, index, category) {
  const col = CALENDAR_BODY_TITLE_COL;
  // Put "category" into the first column
  thisMonthSheet.getRange(row, col - 1).setValue("category");
  // Category heading (A. Work or B. School, etc.)
  thisMonthSheet
    .getRange(row, col)
    .setValue(`${String.fromCharCode(65 + index)}. ${category}`);
  const numColumns = Math.max(
    4, // When creating headers because it's a new month, header rows are created before
      // we make the date columns, so getLastColumn() doesn't return what we need it to...
    thisMonthSheet.getLastColumn() - 1 // The first column is the hidden row-type column,
      // so we minus 1.  The number of columns changes every day as more date columns get added.
  );
  thisMonthSheet.getRange(row, col, 1, numColumns).mergeAcross();
  thisMonthSheet.getRange(row, col)
    .setBackground("black")
    .setFontColor("white")
    .setFontSize(14)
    .setFontWeight("bold");
}

const today = new Date();
function onNewMonth() { prepareNewSheet(today); }
function onNewDay() { prepareNewColumn(today); }
