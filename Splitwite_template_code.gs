/**
 * Splitwise Expense Tracker - Google Apps Script
 * This script creates a form-based expense splitting system in Google Sheets
 */

// Configuration
const MEMBERS_SHEET_NAME = 'Members';
const EXPENSE_SHEET_NAME = 'Expenses';
const SUMMARY_SHEET_NAME = 'Summary';
const FORM_SHEET_NAME = 'Form';

/**
 * Get list of users from Members sheet
 */
function getUsers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const membersSheet = ss.getSheetByName(MEMBERS_SHEET_NAME);
  
  if (!membersSheet) {
    throw new Error('Members sheet not found. Please run Initialize first.');
  }
  
  const lastRow = membersSheet.getLastRow();
  if (lastRow <= 1) {
    throw new Error('No members found in Members sheet. Please add members first.');
  }
  
  const members = [];
  for (let i = 2; i <= lastRow; i++) {
    const member = membersSheet.getRange(i, 1).getValue().toString().trim();
    if (member) {
      members.push(member);
    }
  }
  
  if (members.length === 0) {
    throw new Error('No valid members found. Please add members to the Members sheet.');
  }
  
  console.log('Found members:', members);
  return members;
}

/**
 * Initialize the spreadsheet with required sheets and structure
 */
function initializeSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  owner = ss.getOwner().getEmail()
  currentuser = Session.getActiveUser().getEmail();

  //SpreadsheetApp.getUi().alert(owner);
  //SpreadsheetApp.getUi().alert(currentuser);

  if(owner !== currentuser)
  {
    throw new Error('Only Owners of the sheet can initialized the sheet.');
  }

  // Create or clear sheets
  createOrClearSheet(ss, MEMBERS_SHEET_NAME);
  createOrClearSheet(ss, FORM_SHEET_NAME);
  createOrClearSheet(ss, EXPENSE_SHEET_NAME);
  createOrClearSheet(ss, SUMMARY_SHEET_NAME);
  
  setupMembersSheet();
  setupFormSheet();
  setupExpenseSheet();
  setupSummarySheet();
  
  ensureEditTriggerExists();

  SpreadsheetApp.getUi().alert('Initialization complete! 🎉\n\n📋 Next steps:\n1. Update members in the "Members" sheet\n2. Click "🔄 Refresh Forms" to update dropdowns\n3. Start adding expenses using the form');
}

function ensureEditTriggerExists() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ssId = ss.getId();
  const triggers = ScriptApp.getProjectTriggers();

  const hasEditTrigger = triggers.some(trigger =>
    trigger.getEventType() === ScriptApp.EventType.ON_EDIT &&
    trigger.getHandlerFunction() === 'onEdit' &&
    trigger.getTriggerSourceId() === ssId
  );

  if (!hasEditTrigger) {
    ScriptApp.newTrigger('onEdit')
      .forSpreadsheet(ss)
      .onEdit()
      .create();

    Logger.log(`✅ Created new onEdit trigger for spreadsheet: ${ss.getName()}`);
    ss.toast("✅ onEdit trigger created for this spreadsheet!");
  } else {
    Logger.log(`✔️ onEdit trigger already exists for spreadsheet: ${ss.getName()}`);
    ss.toast("✔️ onEdit trigger already exists.");
  }
}

/**
 * Create or clear a sheet
 */
function createOrClearSheet(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clear();
  }
  return sheet;
}

/**
 * Refresh the From, summay and expense for additional users.
 */
function refreshForms() {
  const USERS = getUsers();
  // Update setupForm
  setupFormSheet();
  // Update Expsnes Form
  setupExpenseSheet();
}

/**
 * Setup the form sheet
 */
function setupFormSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(FORM_SHEET_NAME);
  sheet.clear();
  const USERS = getUsers();
  // Form title
  sheet.getRange('A1').setValue('💰 Expense Splitter Form').setFontSize(16).setFontWeight('bold');
  
  // Form fields
  sheet.getRange('A3').setValue('Expense Details:');
  sheet.getRange('B3').setValue('Enter description here...');
  
  sheet.getRange('A4').setValue('Amount:');
  sheet.getRange('B4').setValue('0');
  
  sheet.getRange('A5').setValue('Date (YYYY-MM-DD):');
  sheet.getRange('B5').setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'));
  
  sheet.getRange('A6').setValue('Paid by:');
  const paidByRange = sheet.getRange('B6');
  const paidByRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(USERS)
    .setAllowInvalid(false)
    .build();
  paidByRange.setDataValidation(paidByRule);
  paidByRange.setValue(USERS[0]);
  
  // Split options
  sheet.getRange('A8').setValue('💡 Split Options:').setFontWeight('bold');
  sheet.getRange('A9').setValue('Equal Split:');
  const equalSplitRange = sheet.getRange('B9');
  const equalSplitRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['FALSE', 'TRUE'])
    .setAllowInvalid(false)
    .build();
  equalSplitRange.setDataValidation(equalSplitRule);
  equalSplitRange.setValue('FALSE');
  
  // Individual split amounts with exclude options
  sheet.getRange('A11').setValue('👥 Individual Split Amounts & Exclude Options:').setFontWeight('bold');
  
  // Headers for the split section
  sheet.getRange('A12').setValue('User');
  sheet.getRange('B12').setValue('Amount');
  sheet.getRange('C12').setValue('Exclude');
  sheet.getRange('A12:C12').setFontWeight('bold').setBackground('#E3F2FD');
  
  let currentRow = 13;
  USERS.forEach(user => {
    sheet.getRange(currentRow, 1).setValue(user);
    sheet.getRange(currentRow, 2).setValue(0);
    
    // Add exclude dropdown
    const excludeRange = sheet.getRange(currentRow, 3);
    const excludeRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['FALSE', 'TRUE'])
      .setAllowInvalid(false)
      .build();
    excludeRange.setDataValidation(excludeRule);
    excludeRange.setValue('FALSE');
    
    currentRow++;
  });
  
  // Submit button (Note: This is just visual - use menu to submit)
  sheet.getRange(currentRow + 1, 1).setValue('🚀 SUBMIT EXPENSE (Use Menu Above)');
  sheet.getRange(currentRow + 1, 1).setBackground('#4CAF50').setFontColor('white').setFontWeight('bold');

  // Add a note about how to submit
  sheet.getRange(currentRow + 2, 1).setValue('⚠️ To submit: Go to menu "💰 Expense Splitter" → "Submit Expense"');
  sheet.getRange(currentRow + 2, 1).setFontColor('red').setFontStyle('italic');

  // Make if mobile friendly.
  sheet.getRange(2, 1).setValue('🚀 SUBMIT EXPENSE use Dropdown -->');
  
  const submitbuttonrange = sheet.getRange(2, 2);
  const submitbuttonRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Submit', 'Reset'])
    .setAllowInvalid(false)
    .build();
  submitbuttonrange.setDataValidation(submitbuttonRule);
  submitbuttonrange.setValue('');

  // Status for mobile users
  sheet.getRange('D1').setValue('Status').setFontWeight('bold');

  // Instructions
  sheet.getRange('E1').setValue('📋 Instructions:').setFontWeight('bold');
  sheet.getRange('E2').setValue('1. Fill expense details and amount');
  sheet.getRange('E3').setValue('2. Select who paid');
  sheet.getRange('E4').setValue('3. Choose equal split OR enter individual amounts');
  sheet.getRange('E5').setValue('4. Click SUBMIT EXPENSE to add');
  sheet.getRange('E6').setValue('5. Wait for status update under Status');
  sheet.getRange('E7').setValue('6. Check Summary sheet for settlements');
  
  // Format columns
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(4, 250);
}

/**
 * Setup expense tracking sheet
 */
function setupExpenseSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(EXPENSE_SHEET_NAME);
  
  // Check if expense sheet alrady has some entries. 
  
  // Get last row with data in column A
  var lastRow = sheet.getRange("A:A").getLastRow();

  if (lastRow > 2) {
    Logger.log("✅ Data exists below A2, not clearing the sheet");
  } else {
    Logger.log("❌ No data below A2");
    sheet.clear();
  }

  // Get current users
  let USERS;
  try {
    USERS = getUsers();
  } catch (error) {
    USERS = ['Member1', 'Member2', 'Member3']; // Fallback
  }
  
  // Headers
  const headers = ['Date', 'Description', 'Total Amount', 'Paid By', ...USERS];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#E3F2FD');
  
  // Summary row for totals
  const summaryRowIndex = 2;
  sheet.getRange(summaryRowIndex, 1).setValue('BALANCE →');
  sheet.getRange(summaryRowIndex, 2).setValue('Net Amount');
  sheet.getRange(summaryRowIndex, 1, 1, 4).setBackground('#FFF3E0').setFontWeight('bold');
  
  // Calculate balance formulas for each user
  for (let i = 0; i < USERS.length; i++) {
    const col = 5 + i; // User columns start at column 5
    const colLetter = String.fromCharCode(64 + col);
    // Formula: Sum of what they paid minus sum of what they owe
    sheet.getRange(summaryRowIndex, col).setFormula(
      `=SUMIF(D:D,"${USERS[i]}",C:C)-SUM(${colLetter}3:${colLetter}1000)`
    );
  }
  
  sheet.getRange(summaryRowIndex, 5, 1, USERS.length).setBackground('#E8F5E8').setFontWeight('bold');
}

/**
 * Setup summary sheet for settlements
 */
function setupSummarySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SUMMARY_SHEET_NAME);
  
  sheet.getRange('A1').setValue('💳 Settlement Summary').setFontSize(16).setFontWeight('bold');
  sheet.getRange('A2').setValue('Who owes money to whom:').setFontStyle('italic');
  
  sheet.getRange('A4').setValue('From');
  sheet.getRange('B4').setValue('To');
  sheet.getRange('C4').setValue('Amount');
  sheet.getRange('A4:C4').setFontWeight('bold').setBackground('#E3F2FD');
}

/**
 * Enhanced submitExpense function with additional validation
 */
function submitExpense() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formSheet = ss.getSheetByName(FORM_SHEET_NAME);
  // Get current users
  const USERS = getUsers();
  
  try {
    
    // Get form data
    const description = formSheet.getRange('B3').getValue().toString().trim();
    const amount = parseFloat(formSheet.getRange('B4').getValue());
    const dateStr = formSheet.getRange('B5').getValue().toString();
    const paidBy = formSheet.getRange('B6').getValue().toString();
    const equalSplitValue = formSheet.getRange('B9').getValue();
    const equalSplitString = equalSplitValue.toString().trim().toUpperCase();
    const equalSplit = equalSplitString === 'TRUE';
    
    // Additional validation to prevent empty submissions
    if (!description || description === 'Enter description here...' || description === '') {
      throw new Error('Please enter expense details');
    }
    
    if (isNaN(amount) || amount <= 0) {
      throw new Error('Please enter a valid amount greater than 0');
    }
    
    // Check for potential duplicate by looking at recent entries
    const expenseSheet = ss.getSheetByName(EXPENSE_SHEET_NAME);
    const lastRow = expenseSheet.getLastRow();
    
    // Check last 3 entries for potential duplicates (same description, amount, and date within 1 minute)
    const currentTime = new Date();
    for (let i = Math.max(3, lastRow - 2); i <= lastRow; i++) {
      if (i <= 2) continue; // Skip header and summary rows
      
      const existingDesc = expenseSheet.getRange(i, 2).getValue().toString();
      const existingAmount = parseFloat(expenseSheet.getRange(i, 3).getValue());
      const existingDate = new Date(expenseSheet.getRange(i, 1).getValue());
      
      // Check if this looks like a duplicate (same description, amount, and within 1 minute)
      const timeDiff = Math.abs(currentTime.getTime() - existingDate.getTime());
      if (existingDesc === description && 
          Math.abs(existingAmount - amount) < 0.01 && 
          timeDiff < 60000) { // 1 minute
        throw new Error('Duplicate expense detected. This expense was already submitted.');
      }
    }
    
    // Validate date
    const date = new Date(dateStr);
    if (isNaN(date.getTime())) {
      throw new Error('Please enter a valid date in YYYY-MM-DD format');
    }
    
    if (!USERS.includes(paidBy)) {
      throw new Error('Please select a valid user who paid');
    }
    
    // Get split amounts
    let splitAmounts = {};
    let totalSplit = 0;
    
    if (equalSplit) {
      // Equal split logic (unchanged)
      console.log('Processing EQUAL split for amount:', amount);
      
      const excludedUsers = [];
      const includedUsers = [];
      
      for (let i = 0; i < USERS.length; i++) {
        const excludeValue = formSheet.getRange(13 + i, 3).getValue().toString().trim().toUpperCase();
        const isExcluded = excludeValue === 'TRUE';
        
        if (isExcluded) {
          excludedUsers.push(USERS[i]);
          splitAmounts[USERS[i]] = 0;
        } else {
          includedUsers.push(USERS[i]);
        }
      }
      
      if (includedUsers.length === 0) {
        throw new Error('Cannot split expense: All users are excluded. At least one person must be included.');
      }
      
      const amountPerPerson = amount / includedUsers.length;
      
      includedUsers.forEach(user => {
        splitAmounts[user] = amountPerPerson;
        totalSplit += amountPerPerson;
      });
      
      totalSplit = amount; // Ensure exact match
      
    } else {
      // Individual amounts logic (unchanged)
      console.log('Processing INDIVIDUAL split amounts');
      for (let i = 0; i < USERS.length; i++) {
        const userAmount = parseFloat(formSheet.getRange(13 + i, 2).getValue()) || 0;
        splitAmounts[USERS[i]] = userAmount;
        totalSplit += userAmount;
      }
      
      if (Math.abs(totalSplit - amount) > 0.01) {
        throw new Error(`Split amounts (${totalSplit.toFixed(2)}) don't match expense amount (${amount.toFixed(2)}). Please enter individual amounts that add up to the total expense.`);
      }
    }
    
    // Add to expense sheet
    const newRow = lastRow + 1;
    
    // Insert the expense data
    const rowData = [
      Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      description,
      amount,
      paidBy
    ];
    
    // Add split amounts for each user
    USERS.forEach(user => {
      rowData.push(splitAmounts[user]);
    });
    
    expenseSheet.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);
    
    // Update summary
    updateSummary();
    
    // Clear form
    clearForm();
    
    console.log('Expense submitted successfully:', description, amount);
    
  } catch (error) {
    console.error('Error in submitExpense:', error);
    throw error;
  }
}

/**
 * Setup the members sheet
 */
function setupMembersSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(MEMBERS_SHEET_NAME);
  
  // Header
  sheet.getRange('A1').setValue('Members').setFontWeight('bold').setBackground('#E3F2FD');
  
  // Sample data
  const sampleMembers = ['Alice', 'Bob', 'Charlie', 'Diana', 'Eve'];
  for (let i = 0; i < sampleMembers.length; i++) {
    sheet.getRange(i + 2, 1).setValue(sampleMembers[i]);
  }
  
  // Instructions
  sheet.getRange('C1').setValue('📋 Instructions:').setFontWeight('bold');
  sheet.getRange('C2').setValue('• Add/remove member names in column A');
  sheet.getRange('C3').setValue('• Each member should be in a separate row');
  sheet.getRange('C4').setValue('• After changing members, refresh forms');
  sheet.getRange('C5').setValue('• Use menu: "💰 Expense Splitter" → "🔄 Refresh Forms"');
  sheet.getRange('C6').setValue('');
  sheet.getRange('C7').setValue('💡 Tips:').setFontWeight('bold').setFontColor('blue');
  sheet.getRange('C8').setValue('• Members listed here will appear in expense forms');
  sheet.getRange('C9').setValue('• Only people with Editor access can submit expenses');
  sheet.getRange('C10').setValue('• You can have members without edit access for splitting');
  
  // Format
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(3, 350);
}
/**
 * Enhanced clearForm function
 */
function clearForm() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(FORM_SHEET_NAME);
  const USERS = getUsers();
  sheet.getRange('B3').setValue('Enter description here...');
  sheet.getRange('B4').setValue(0);
  sheet.getRange('B5').setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'));
  sheet.getRange('B9').setValue('FALSE');

  // Clear status and lock cells
  sheet.getRange('D2').clear();
  sheet.getRange('D2').setBackground('white');
  sheet.getRange('D3').clear(); // Clear lock cell
  sheet.getRange('D3').setBackground('white');
  
  // Clear individual amounts and exclude flags
  for (let i = 0; i < USERS.length; i++) {
    sheet.getRange(13 + i, 2).setValue(0); // Amount
    sheet.getRange(13 + i, 3).setValue('FALSE'); // Exclude
  }
}

/**
 * Update settlement summary
 */
function updateSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const expenseSheet = ss.getSheetByName(EXPENSE_SHEET_NAME);
  const summarySheet = ss.getSheetByName(SUMMARY_SHEET_NAME);
  
  const USERS = getUsers();

  // Get balances from expense sheet
  const balances = {};
  for (let i = 0; i < USERS.length; i++) {
    const col = 5 + i;
    const balance = expenseSheet.getRange(2, col).getValue();
    balances[USERS[i]] = parseFloat(balance) || 0;
  }
  
  // Calculate settlements using a simple algorithm
  const settlements = calculateSettlements(balances);
  
  // Clear existing settlement data
  summarySheet.getRange('A5:C1000').clear();
  
  // Write settlements
  let row = 5;
  settlements.forEach(settlement => {
    summarySheet.getRange(row, 1).setValue(settlement.from);
    summarySheet.getRange(row, 2).setValue(settlement.to);
    summarySheet.getRange(row, 3).setValue(settlement.amount.toFixed(2));
    row++;
  });
  
  if (settlements.length === 0) {
    summarySheet.getRange('A5').setValue('🎉 All settled up!');
  }
}

/**
 * Calculate optimal settlements
 */
function calculateSettlements(balances) {
  const settlements = [];
  const creditors = []; // People who are owed money (positive balance)
  const debtors = [];   // People who owe money (negative balance)
  
  // Separate creditors and debtors
  Object.keys(balances).forEach(user => {
    const balance = balances[user];
    if (balance > 0.01) {
      creditors.push({ user, amount: balance });
    } else if (balance < -0.01) {
      debtors.push({ user, amount: -balance });
    }
  });
  
  // Sort by amount (largest first)
  creditors.sort((a, b) => b.amount - a.amount);
  debtors.sort((a, b) => b.amount - a.amount);
  
  // Match creditors with debtors
  let i = 0, j = 0;
  while (i < creditors.length && j < debtors.length) {
    const creditor = creditors[i];
    const debtor = debtors[j];
    
    const settleAmount = Math.min(creditor.amount, debtor.amount);
    
    if (settleAmount > 0.01) {
      settlements.push({
        from: debtor.user,
        to: creditor.user,
        amount: settleAmount
      });
      
      creditor.amount -= settleAmount;
      debtor.amount -= settleAmount;
    }
    
    if (creditor.amount < 0.01) i++;
    if (debtor.amount < 0.01) j++;
  }
  
  return settlements;
}

/**
 * Install menu and triggers
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  try{
    ui.createMenu('💰 Expense Splitter')
      .addItem('🔧 Initialize Sheets', 'initializeSpreadsheet')
      .addItem('➕ Submit Expense', 'submitExpense')
      .addItem('🔄 Refresh Summary', 'updateSummary')
      .addItem('🔄 Refresh Forms', 'refreshForms')
      .addSeparator()
      .addItem('❓ Help', 'showHelp')
      .addToUi();
  } catch (error) {}
}
/**
 * Support form buttons
 */

/**
 * Fixed onEdit function with duplicate prevention
 */
function onEdit(e) {
  const sheet = e.range.getSheet();
  const cell = e.range;
  
  // Only act on the Form sheet, cell B2
  if (sheet.getName() !== 'Form' || cell.getA1Notation() !== 'B2') return;

  const action = cell.getValue().toString().trim().toUpperCase();
  const statusCell = sheet.getRange("D2");
  
  // Immediately clear the cell to prevent duplicate triggers
  if (action === 'SUBMIT' || action === 'RESET') {
    cell.setValue('');
  }
  
  // Check if we're already processing (simple lock mechanism)
  const lockCell = sheet.getRange("D3");
  const currentLock = lockCell.getValue().toString();
  
  if (currentLock === 'PROCESSING') {
    console.log('Already processing, skipping duplicate execution');
    return;
  }

  try {
    if (action === 'SUBMIT') {
      // Set lock
      lockCell.setValue('PROCESSING');
      lockCell.setBackground('#FFE082');
      
      statusCell.setValue("⏳ Processing...");
      statusCell.setBackground('#FFE082').setFontColor('black').setFontWeight('bold');
      
      // Add a small delay to ensure UI updates
      Utilities.sleep(100);
      
      submitExpense();
      
      statusCell.setValue("✅ Expense Submitted");
      statusCell.setBackground('#4CAF50').setFontColor('white').setFontWeight('bold');

    } else if (action === 'RESET') {
      // Set lock
      lockCell.setValue('PROCESSING');
      lockCell.setBackground('#FFE082');
      
      clearForm();
      
      statusCell.setValue("🧹 Form Cleared");
      statusCell.setBackground('#2196F3').setFontColor('white').setFontWeight('bold');
    }
  } catch (error) {
    console.error('Error in onEdit:', error);
    statusCell.setValue("❌ Error: " + error.message);
    statusCell.setBackground('#F44336').setFontColor('white').setFontWeight('bold');
  } finally {
    // Always clear the lock
    lockCell.setValue('');
    lockCell.setBackground('white');
  }
}

/**
 * Show help dialog
 */
function showHelp() {
  const helpText = `
Splitwise Expense Tracker Help:

1. SETUP:
   - Run "Initialize Sheets" from the menu first
   - Update USERS array in script with your group members

2. ADDING EXPENSES:
   - Use the Form sheet to add new expenses
   - Fill all required fields
   - Choose equal split OR enter individual amounts
   - Click "Submit Expense" from menu or use submitExpense()

3. VIEWING RESULTS:
   - Expenses sheet shows all transactions and balances
   - Summary sheet shows who owes whom
   - Green row shows net balance for each person

4. PERMISSIONS:
   - Only people with edit access can add expenses
   - View-only users can see balances and summaries

Need more help? Check the script comments or contact your sheet admin.
  `;
  
  SpreadsheetApp.getUi().alert('Help', helpText, SpreadsheetApp.getUi().ButtonSet.OK);
}
