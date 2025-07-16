/**
 * @file Code.gs
 * @description This script provides the core functionality for the PROSPR financial planning template,
 * including a custom Admin menu with access control and a Monthly Comparative Report tool.
 * This file is designed to be a self-contained Google Apps Script project for submission.
 *
 * It adheres to principles of clean code and modularity, with extensive comments.
 */

/**
 * The onOpen function is a special Google Apps Script trigger that runs automatically
 * when a user opens the spreadsheet. It's used here to initialize the custom 'Admin' menu.
 */
function onOpen() {
  // Add the Admin menu with password protection when the spreadsheet is opened.
  addAdminMenuWithAccessControl();
  // Note: Original script included a ProsprScript.onOpen() call.
  // For this self-contained submission, it's assumed that any base ProsprScript
  // functionality is either integrated or not required for this specific task.
}

/**
 * Adds a custom 'Admin' menu to the Google Sheet UI.
 * Initially, this menu only contains an 'Unlock Admin Menu' option.
 */
function addAdminMenuWithAccessControl() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Admin')
    .addItem('Unlock Admin Menu', 'showAdminPrompt')
    .addToUi();
}

/**
 * Displays a prompt for the admin code. If the code is correct, it unlocks the full Admin menu.
 */
function showAdminPrompt() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Admin Access', 'Please enter the admin code to unlock admin options:', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() === ui.Button.OK) {
    var code = response.getResponseText();
    if (verifyAdminCode(code)) {
      addUnlockedAdminMenu();
      ui.alert('Admin options unlocked!');
    } else {
      ui.alert('Incorrect code. Access denied.');
    }
  }
}

/**
 * Verifies the admin code against the user's stored property.
 * If no code is set, it will prompt to set the initial code 'PROSPR2025'.
 * This uses UserProperties, ensuring each user (Google account) interacting with the script
 * has their own distinct admin code, which is ideal for multi-client scenarios.
 * @param {string} code The code entered by the user.
 * @returns {boolean} True if the code is correct, false otherwise.
 */
function verifyAdminCode(code) {
  var userProperties = PropertiesService.getUserProperties();
  var storedCode = userProperties.getProperty('ADMIN_CODE');

  if (!storedCode) {
    // If no admin code is set for the current user, set the initial one.
    // This happens on the first attempt to unlock the menu by a new user.
    userProperties.setProperty('ADMIN_CODE', 'PROSPR2025');
    storedCode = 'PROSPR2025';
    SpreadsheetApp.getUi().alert('Initial admin code "PROSPR2025" has been set for your user account. Please try unlocking the menu again.');
    return false; // Initial setup, user needs to re-enter the code.
  }

  return code === storedCode;
}

/**
 * Adds the full 'Admin' menu with all options (report, set code, lock menu).
 * This is called after successful admin code verification.
 */
function addUnlockedAdminMenu() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Admin')
    .addItem('Monthly Comparative Report', 'runMonthlyComparativeReport')
    .addItem('Set Admin Code', 'setNewAdminCode') // New menu item to change admin code
    .addItem('Lock Admin Menu', 'resetAdminMenu')
    .addToUi();
}

/**
 * Resets the Admin menu back to its locked state (only 'Unlock Admin Menu' visible).
 */
function resetAdminMenu() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Admin')
    .addItem('Unlock Admin Menu', 'showAdminPrompt')
    .addToUi();
  ui.alert('Admin menu locked again.');
}

/**
 * Prompts the user to set a new admin code. Requires current code verification for security.
 */
function setNewAdminCode() {
  var ui = SpreadsheetApp.getUi();
  var userProperties = PropertiesService.getUserProperties();

  var currentCodeResponse = ui.prompt('Verify Current Admin Code', 'Please enter your current admin code to change it:', ui.ButtonSet.OK_CANCEL);
  if (currentCodeResponse.getSelectedButton() === ui.Button.OK) {
    var currentCode = currentCodeResponse.getResponseText();
    if (verifyAdminCode(currentCode)) { // Use verifyAdminCode to check current code
      var newCodeResponse = ui.prompt('Set New Admin Code', 'Enter the new admin code:', ui.ButtonSet.OK_CANCEL);
      if (newCodeResponse.getSelectedButton() === ui.Button.OK) {
        var newCode = newCodeResponse.getResponseText();
        if (newCode) {
          userProperties.setProperty('ADMIN_CODE', newCode);
          ui.alert('Admin code successfully updated!');
        } else {
          ui.alert('New admin code cannot be empty.');
        }
      }
    } else {
      ui.alert('Incorrect current admin code. Cannot change.');
    }
  }
}

/**
 * Generates the Monthly Comparative Report based on the 'Monthly Budget' tab.
 * It processes budget data, identifies deviations, creates a new report sheet,
 * and sends a draft email summary.
 */
function runMonthlyComparativeReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var budgetSheetName = 'Monthly Budget';
  var sheet = ss.getSheetByName(budgetSheetName);
    
  // Retrieve report period details from the sheet's designated cells.
  var year = sheet.getRange('F2').getValue();
  var month = sheet.getRange('F3').getValue();
  var bom = sheet.getRange('H2').getValue(); // Beginning of Month date
  var eom = sheet.getRange('H3').getValue(); // End of Month date

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Error: Sheet "' + budgetSheetName + '" not found! Please ensure the "Monthly Budget" tab exists.');
    return;
  }

  // Define column indices for data extraction (0-indexed for arrays).
  var CATEGORY_COL = 1;    // Column B for category headers (e.g., "Shelter")
  var ITEM_DESC_COL = 2;   // Column C for item descriptions (e.g., "Mortgage")
  var BUDGET_COL = 3;      // Column D for Budgeted amounts
  var ACTUAL_COL = 5;      // Column F for Actual amounts

  // --- Configuration ---
  var DEVIATION_THRESHOLD = 0.20; // 20% threshold for reporting significant deviation (e.g., 0.20 for 20%)
  var START_ROW = 5;             // Data parsing starts from row 5 in the 'Monthly Budget' sheet.

  var data = sheet.getDataRange().getValues(); // Get all data from the active sheet.

  // Objects to store parsed data:
  // `allCategoriesData` will hold aggregated data for each main category.
  var allCategoriesData = {};
  var currentCategoryName = null;
  var currentCategoryItems = [];
  var currentCategoryTotalBudget = 0;
  var currentCategoryTotalActual = 0;

  // Loop through each row of the budget data to parse categories, items, and their totals.
  for (var i = START_ROW - 1; i < data.length; i++) {
    var row = data[i];
    // Safely convert cell values to string and trim whitespace.
    var categoryHeader = (row[CATEGORY_COL] !== null && row[CATEGORY_COL] !== undefined) ? String(row[CATEGORY_COL]).trim() : "";
    var itemDescription = (row[ITEM_DESC_COL] !== null && row[ITEM_DESC_COL] !== undefined) ? String(row[ITEM_DESC_COL]).trim() : "";
    
    // Parse budget and actual values, treating empty or non-numeric values as 0.
    var budgetValue = parseFloat(row[BUDGET_COL]) || 0;
    var actualValue = parseFloat(row[ACTUAL_COL]) || 0;

    // Detect a new main category header (e.g., "Shelter", "Food & Supplies").
    // Ensure it's not empty and not a "Total" row (which marks the end of a category block).
    if (categoryHeader && categoryHeader.indexOf('Total') === -1) {
      // If we were processing a previous category, save its aggregated data before starting a new one.
      if (currentCategoryName) {
        allCategoriesData[currentCategoryName] = {
          items: currentCategoryItems,
          totalBudget: currentCategoryTotalBudget,
          totalActual: currentCategoryTotalActual
        };
      }
      // Initialize variables for the new category.
      currentCategoryName = categoryHeader;
      currentCategoryItems = [];
      currentCategoryTotalBudget = 0;
      currentCategoryTotalActual = 0;
    }

    // Add item details to the current category.
    // This condition ensures that:
    // 1. A category is currently being processed (`currentCategoryName`).
    // 2. The row has an item description OR non-zero budget/actual values (to capture items with only one value).
    // 3. The row is NOT a "Total" row (this prevents the "Total" line from being duplicated as an item).
    if (currentCategoryName && (itemDescription || budgetValue !== 0 || actualValue !== 0) && categoryHeader.indexOf('Total') === -1) {
        currentCategoryItems.push({
          description: itemDescription,
          budget: budgetValue,
          actual: actualValue
        });
    }

    // Detect a "Total" row, which signifies the end of a category's data block.
    if (categoryHeader.indexOf('Total') === 0 && currentCategoryName) {
      // Assign the total budget and actual values for the current category.
      currentCategoryTotalBudget = budgetValue;
      currentCategoryTotalActual = actualValue;

      // Save the complete category data (items and totals) to `allCategoriesData`.
      allCategoriesData[currentCategoryName] = {
        items: currentCategoryItems,
        totalBudget: currentCategoryTotalBudget,
        totalActual: currentCategoryTotalActual
      };

      // Reset variables to prepare for the next category.
      currentCategoryName = null;
      currentCategoryItems = [];
      currentCategoryTotalBudget = 0;
      currentCategoryTotalActual = 0;
    }
  }

  // Handle the last category in the sheet if the loop ends without encountering its "Total" row.
  if (currentCategoryName) {
    allCategoriesData[currentCategoryName] = {
      items: currentCategoryItems,
      totalBudget: currentCategoryTotalBudget,
      totalActual: currentCategoryTotalActual
    };
  }

  // Prepare data structures for the tabular report sheet and the email draft.
  var reportRowsForSheet = [
    [
      "Category",         // Column A header
      "Item Description", // Column B header
      "Actual",           // Column C header
      "Planned",          // Column D header (mapped from Budget)
      "Deviation ($)",    // Column E header
      "Deviation (%)",    // Column F header
      "Status"            // Column G header
    ]
  ];
  var reportLinesForEmail = []; // Array to build the email body content.

  // Iterate through all parsed categories to build the detailed report.
  for (var catName in allCategoriesData) {
    var d = allCategoriesData[catName];
    var actual = d.totalActual || 0;
    var budget = d.totalBudget || 0;
    
    // Calculate deviation in absolute terms.
    var deviation = actual - budget; 
    
    // Calculate percentage deviation, handling division by zero for zero budgets.
    var deviationPct = 0;
    if (budget === 0) {
      deviationPct = (actual === 0) ? 0 : (actual > 0 ? 1 : -1); // If budget is 0, actual > 0 means 100% over, actual < 0 means 100% under.
    } else {
      deviationPct = deviation / budget;
    }

    var deviationPctStr = (deviationPct * 100).toFixed(1) + "%";
    // Determine status based on deviation threshold.
    var status = Math.abs(deviationPct) > DEVIATION_THRESHOLD
      ? (deviationPct > 0 ? "Over" : "Under")
      : "OK";
    var deviationSign = deviation > 0 ? "+" : ""; // Add '+' sign for positive deviations.
    var deviationStr = deviationSign + deviation.toFixed(2);

    // Include categories in the report only if they have significant deviation
    // or if there's a value in one column but zero in the other (indicating a notable difference).
    if (status !== "OK" || (budget === 0 && actual !== 0) || (budget !== 0 && actual === 0)) {
      // Add the category summary row for the Google Sheet report.
      reportRowsForSheet.push([
        catName, // Category name in Column A
        "",      // Empty for Column B (Item Description)
        "$" + actual.toFixed(2), // Formatted Actual value
        "$" + budget.toFixed(2), // Formatted Planned value
        deviationStr,
        deviationPctStr,
        status
      ]);

      // Add the category summary line for the email report.
      reportLinesForEmail.push(
        catName + ": " + status + " budget by " + deviationPctStr +
        " ($" + actual.toFixed(2) + " vs. $" + budget.toFixed(2) + ")"
      );

      // Collect and report on significant individual items within this category.
      var significantItems = [];
      for (var j = 0; j < d.items.length; j++) {
        var item = d.items[j];
        var itemDeviation = item.actual - item.budget;
        var itemDeviationPct = 0;

        if (item.budget === 0) {
          itemDeviationPct = (item.actual === 0) ? 0 : (item.actual > 0 ? 1 : -1);
        } else {
          itemDeviationPct = itemDeviation / item.budget;
        }

        // Highlight items if their percentage deviation exceeds the threshold,
        // or if one value is zero and the other is not.
        if (Math.abs(itemDeviationPct) > DEVIATION_THRESHOLD || (item.budget === 0 && item.actual !== 0) || (item.budget !== 0 && item.actual === 0)) {
          var itemDiffSign = itemDeviation > 0 ? "+" : "";
          significantItems.push({
            description: item.description,
            actual: item.actual,
            budget: item.budget,
            diff: itemDiffSign + itemDeviation.toFixed(2),
            diffPct: (itemDeviationPct * 100).toFixed(1) + "%"
          });

          // Add detailed item row for the sheet (indented in Column B).
          reportRowsForSheet.push([
            "", // Empty for Column A (Category)
            item.description, // Item description in Column B
            "$" + item.actual.toFixed(2),
            "$" + item.budget.toFixed(2),
            itemDiffSign + itemDeviation.toFixed(2),
            (itemDeviationPct * 100).toFixed(1) + "%", // Show % deviation for individual items
            ""  // No status for individual items
          ]);
        }
      }

      // Add key items to the email summary if any significant items were found.
      if (significantItems.length > 0) {
        reportLinesForEmail.push("   Key Items:");
        significantItems.forEach(function(item) {
          reportLinesForEmail.push(
            "     " + item.description + ": $" + item.actual.toFixed(2) + " (Actual) vs $" + item.budget.toFixed(2) + " (Planned) (Diff: " + item.diff + ", " + item.diffPct + ")"
          );
        });
      }
      reportLinesForEmail.push(""); // Add a blank line between categories for email clarity.
        
      // Add an empty row after each category block for visual separation in the sheet.
      reportRowsForSheet.push(["", "", "", "", "", "", ""]);  
    }
  }

  // Generate the tabular report in a new sheet.
  generateTabularReportSheet(reportRowsForSheet, month, year, bom, eom);

  // Prepare and generate the email draft summary.
  var bomStr = prettyDate(bom);
  var eomStr = prettyDate(eom);
  var finalReportText =
    'Monthly Budget Deviation Report\n' +
    'Period: ' + month + ' ' + year + ' (' + bomStr + ' - ' + eomStr + ')\n' +
    'Generated on: ' + new Date().toLocaleDateString() + '\n\n' +
    reportLinesForEmail.join("\n");
  
  var emailSubject = month + ' ' + year + ' Budget Comparison';
  generateReportAsEmailDraft(finalReportText, emailSubject);
}

/**
 * Formats a Date object into a "MM/dd/yyyy" string.
 * @param {Date} date The date object to format.
 * @returns {string} The formatted date string.
 */
function prettyDate(date) {
  // Ensure the input is a Date object before formatting.
  if (!(date instanceof Date)) return date;
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "MM/dd/yyyy");
}

/**
 * Generates a tabular report in a new Google Sheet.
 * This function handles sheet creation/clearing, header population,
 * data insertion, and formatting.
 * @param {Array<Array<any>>} tableData The data for the main report table.
 * @param {string} month The month for the report.
 * @param {number} year The year for the report.
 * @param {Date} bom The beginning of the month date.
 * @param {Date} eom The end of the month date.
 */
function generateTabularReportSheet(tableData, month, year, bom, eom) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var reportSheetName = month + ' Budget Comparison'; 
  var reportSheet = ss.getSheetByName(reportSheetName);
  
  // Clear or create the report sheet.
  if (reportSheet) {
    reportSheet.clear(); // Clear existing content if sheet exists.
  } else {
    reportSheet = ss.insertSheet(reportSheetName); // Create new sheet if it doesn't exist.
  }

  // Populate the top-right header section with report period details.
  reportSheet.getRange('C2').setValue("Year");
  reportSheet.getRange('D2').setValue(year);
  reportSheet.getRange('E2').setValue("BOM");
  reportSheet.getRange('F2').setValue(Utilities.formatDate(bom, Session.getScriptTimeZone(), "M/d/yyyy"));

  reportSheet.getRange('C3').setValue("Month");
  reportSheet.getRange('D3').setValue(month);
  reportSheet.getRange('E3').setValue("EOM");
  reportSheet.getRange('F3').setValue(Utilities.formatDate(eom, Session.getScriptTimeZone(), "M/d/yyyy"));

  // Output the main report table data starting at row 5, column 1.
  var nRows = tableData.length;
  var nCols = 7; // Number of columns in the report table.
  var dataStartRow = 5; // The row where the main table data begins.
  reportSheet.getRange(dataStartRow, 1, nRows, nCols).setValues(tableData);

  // Apply formatting to the header row of the report table.
  var headerRange = reportSheet.getRange(dataStartRow, 1, 1, nCols);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#e3e3e3'); // Light grey background.
  headerRange.setFontFamily('Arial, Helvetica, sans-serif');

  // Apply a consistent font family to all data rows below the header.
  if (nRows > 1) { // Ensure there are actual data rows.
    reportSheet.getRange(dataStartRow + 1, 1, nRows - 1, nCols).setFontFamily('Arial, Helvetica, sans-serif');
  }

  // Apply conditional formatting and styling to data rows based on content.
  for (var i = 0; i < nRows - 1; i++) { // Loop through data rows (skip the table header row).
    var currentRowIndex = dataStartRow + 1 + i; // Actual row index in the sheet.
    var rowData = tableData[i + 1]; // Get the corresponding data from the `tableData` array.

    var categoryCellContent = String(rowData[0]); // Content of Column A for the current row.

    // If the row is an empty separator row (added for visual spacing).
    if (categoryCellContent === "") {
        reportSheet.getRange(currentRowIndex, 1, 1, nCols).setBackground('#ffffff'); // White background.
        continue; // Skip further formatting for empty rows.
    }

    // If Column A is empty but Column B has content, it's an indented item row.
    if (rowData[0] === "" && rowData[1] !== "") { 
      reportSheet.getRange(currentRowIndex, 1, 1, nCols).setBackground('#f7f7f7'); // Light grey background.
      reportSheet.getRange(currentRowIndex, 2).setFontSize(9); // Smaller font for item description.
      reportSheet.getRange(currentRowIndex, 2).setHorizontalAlignment('left'); // Left align item description.
    } else {
      // Otherwise, it's a main category summary row.
      var status = rowData[6]; // Get the 'Status' from Column G.
      var statusCell = reportSheet.getRange(currentRowIndex, 7); // Reference to the Status cell.
      
      // Apply background color based on status (Over/Under budget).
      if (status === "Over") {
        statusCell.setBackground('#ffd6d6'); // Light red for over budget.
      } else if (status === "Under") {
        statusCell.setBackground('#d6ffd6'); // Light green for under budget.
      }
      reportSheet.getRange(currentRowIndex, 1).setFontWeight('bold'); // Bold category name.
      reportSheet.getRange(currentRowIndex, 1).setFontLine('underline'); // Underline category name.
      reportSheet.getRange(currentRowIndex, 1, 1, nCols).setBackground('#f0f8ff'); // Light blue background for categories.
    }
  }

  // Auto-resize all columns for optimal readability.
  for (var c = 1; c <= nCols; c++) {
    reportSheet.autoResizeColumn(c);
  }

  // Right-align numeric columns for better presentation.
  reportSheet.getRange(dataStartRow, 3, nRows, 1).setHorizontalAlignment('right'); // Actual (Column C)
  reportSheet.getRange(dataStartRow, 4, nRows, 1).setHorizontalAlignment('right'); // Planned (Column D)
  reportSheet.getRange(dataStartRow, 5, nRows, 1).setHorizontalAlignment('right'); // Deviation ($) (Column E)
  reportSheet.getRange(dataStartRow, 6, nRows, 1).setHorizontalAlignment('right'); // Deviation (%) (Column F)

  // Set the newly generated report sheet as the active sheet for user visibility.
  ss.setActiveSheet(reportSheet);
  SpreadsheetApp.getUi().alert('Report generated successfully in the "' + reportSheetName + '" sheet.');
}

/**
 * Creates a draft email in Gmail with the report content.
 * This provides an alternative output format for the report summary.
 * @param {string} content The plain text content of the report summary for the email body.
 * @param {string} subject The subject line for the email draft.
 */
function generateReportAsEmailDraft(content, subject) {
  try {
    // Create a new Gmail draft. The content is wrapped in <pre> tags for monospace formatting.
    GmailApp.createDraft('', subject, '', { htmlBody: '<pre>' + content + '</pre>' });
    SpreadsheetApp.getUi().alert('Report has been saved as a draft in your Gmail.');
  } catch (e) {
    // Catch block to handle potential Gmail permissions issues.
    SpreadsheetApp.getUi().alert('Could not create Gmail draft. Please ensure you have granted script permissions for Gmail. Error: ' + e.message);
  }
}