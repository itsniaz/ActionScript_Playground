//Phase 1
function onInstall(e) {
    onOpen(e);// Call onOpen to create menu items right after installation.
  }
  
  function onOpen(e) {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Custom Menu')
      .addItem('Statement', 'mainFunction')
      .addItem('Nuvei-Journal', 'mainFunctionA')
      .addItem('Nuvei-Splittabs', 'mainFunctionB')
      .addItem('Nuvei-Customertab', 'mainFunctionC')
      .addItem('Nuvei-Finetune', 'mainFunctionD')
      .addToUi();
  }
  
  function mainFunction(){
    runAll1(); //for the creation of Statement
    runAll2();//for fetching data from revenue sheet
  }
  function runAll1() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const sheetName = "Statement";
  
    // Step 1: Setup Statement Tab
    const existingTab = ss.getSheetByName(sheetName);
    if (existingTab) {
      ss.deleteSheet(existingTab);
    }
    const statementSheet = ss.insertSheet(sheetName);
    const headers = [
      "System Generated Order", "Merchant Defined Order Number", "Payment Status", 
      "Date", "Total", "Card Holder Email", "Refund Amount", "Refund Date", 
      "Order Status", "Custom Data"
    ];
    statementSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setBackground("green").setFontColor("yellow");
  
    // Step 2: Read Source Sheet Data
    const sourceData = sheet.getDataRange().getValues();
    const sourceOrderIdIndex = 0;  // Column A (System Generated Order)
    const sourceMerchantOrderIndex = 1;  // Column B (Merchant Defined Order Number)
    const sourceOrderStatusIndex = 2;  // Column C (Order Status)
    const sourcePaymentStatusIndex = 3;  // Column D (Payment Status)
    const sourceDateIndex = 6;  // Column G (Date)
    const sourceTotalIndex = 8;  // Column I (Total)
    const sourceEmailIndex = 13;  // Column N (Card Holder Email)
    const sourceRefundAmountIndex = 36;  // Column AK (Refund Amount)
    const sourceRefundDateIndex = 37;  // Column AL (Refund Date)
  
    const sourceMap = [];
    for (let i = 1; i < sourceData.length; i++) {
      const orderId = sourceData[i][sourceOrderIdIndex];
      const merchantOrderId = sourceData[i][sourceMerchantOrderIndex];
      const orderStatus = sourceData[i][sourceOrderStatusIndex];  // Order Status
      const paymentStatus = sourceData[i][sourcePaymentStatusIndex];  // Payment Status
      const date = sourceData[i][sourceDateIndex];
      const total = sourceData[i][sourceTotalIndex];
      const email = sourceData[i][sourceEmailIndex];
      const refundAmount = sourceData[i][sourceRefundAmountIndex];
      const refundDate = sourceData[i][sourceRefundDateIndex];
  
      // **Filter for SUCCESS transactions only**
      if (paymentStatus === "SUCCESS") {
        sourceMap.push([
          orderId, merchantOrderId, paymentStatus, date, total, email, refundAmount, refundDate, orderStatus
        ]);
      }
    }
  
    // Step 3: Prompt for External Sheet and Tab
    const externalLink = Browser.inputBox("Enter the link for the external sheet:");
    if (!externalLink) {
      SpreadsheetApp.getUi().alert("No link provided. Exiting.");
      return;
    }
  
    let externalSpreadsheet;
    try {
      externalSpreadsheet = SpreadsheetApp.openByUrl(externalLink);
    } catch (e) {
      SpreadsheetApp.getUi().alert("Could not open the external sheet. Please check the link.");
      return;
    }
  
    const externalTabName = Browser.inputBox("Enter the tab name in the external sheet:");
    const externalSheet = externalSpreadsheet.getSheetByName(externalTabName);
    if (!externalSheet) {
      SpreadsheetApp.getUi().alert("The specified tab was not found in the external sheet.");
      return;
    }
  
    const externalData = externalSheet.getDataRange().getValues();
    const externalHeaderRow = externalData[0];
    const transactionIdIndex = externalHeaderRow.indexOf("Transaction ID"); // Column B
    const customDataIndex = externalHeaderRow.indexOf("Custom Data"); // Column U
  
    if (
      transactionIdIndex === -1 ||
      customDataIndex === -1
    ) {
      SpreadsheetApp.getUi().alert("Required columns were not found in the external sheet.");
      return;
    }
  
    // Step 4: Match and Pull Data
    const externalMap = {};
    for (let i = 1; i < externalData.length; i++) {
      const transactionId = externalData[i][transactionIdIndex];
      const customData = externalData[i][customDataIndex];
  
      if (transactionId) {
        externalMap[transactionId] = customData;
      }
    }
  
    // Step 5: Populate Statement Tab
    const outputData = [];
    for (let i = 0; i < sourceMap.length; i++) {
      const row = sourceMap[i];
      const merchantOrderId = row[1];  // Merchant Defined Order Number
      const customData = externalMap[merchantOrderId] || "";  // Blank if no match found
      row.push(customData);  // Add Custom Data or blank to the end of the row
      outputData.push(row);
    }
  
    // Write the final data to the Statement Tab
    if (outputData.length > 0) {
      statementSheet.getRange(2, 1, outputData.length, headers.length).setValues(outputData);
    }
  
    SpreadsheetApp.getUi().alert("Statement tab has been successfully populated.");
  }
  // Phase 2: runAll2 - Fetch External Sheet Data into Columns K-Z
  function runAll2() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const localSheet = ss.getSheetByName("Statement");
    if (!localSheet) {
      throw new Error("Local sheet 'Statement' not found.");
    }
  
    const localData = localSheet.getDataRange().getValues();
    const customDataIndex = 9; // Column J (Custom Data)
  
    // Step 1: Add headers for columns K to Z
    const headers = [
      "payment_method", "payment_date", "transaction_id", "name", "email", 
      "package", "account_type", "order_type", "revenue_before_discount", 
      "discount", "coupon_code", "add_on_amount", "add_on_title", "net_revenue", 
      "commission", "order_id"
    ];
    localSheet.getRange(1, 11, 1, headers.length).setValues([headers]).setBackground("green").setFontColor("yellow");
  
    const externalSheetLinks = [];
  
    // Step 2: Prompt for External Sheet Links
    for (let i = 0; i < 2; i++) {
      const link = Browser.inputBox(`Enter the link for external sheet ${i + 1} (or 0 to skip):`);
      if (link === "0") {
        continue;
      }
      if (!link) {
        SpreadsheetApp.getUi().alert(`No link provided for sheet ${i + 1}. Exiting.`);
        return;
      }
  
      const tabNames = [];
      for (let j = 0; j < 3; j++) {
        const tabName = Browser.inputBox(`Enter the name for tab ${j + 1} in sheet ${i + 1} (or 0 to skip):`);
        if (tabName === "0") {
          continue;
        }
        tabNames.push(tabName);
      }
  
      externalSheetLinks.push({ link, tabNames });
    }
  
    if (externalSheetLinks.length === 0) {
      SpreadsheetApp.getUi().alert("No valid links provided. Exiting.");
      return;
    }
  
    // Step 3: Process Each External Sheet and Tab
    externalSheetLinks.forEach(({ link, tabNames }) => {
      try {
        const externalSpreadsheet = SpreadsheetApp.openByUrl(link);
  
        tabNames.forEach(tabName => {
          const externalSheet = externalSpreadsheet.getSheetByName(tabName);
          if (!externalSheet) {
            SpreadsheetApp.getUi().alert(`Tab '${tabName}' not found in sheet '${link}'.`);
            Logger.log(`Tab '${tabName}' not found. Skipping.`);
            return;
          }
  
          const externalData = externalSheet.getDataRange().getValues();
          const externalHeaderRow = externalData[0];
          const transactionIdIndex = externalHeaderRow.indexOf("transaction_id");
  
          // Mapping columns for extraction
          const columnMap = {
            payment_method: 1, // Column B
            payment_date: 2,   // Column C
            transaction_id: 4, // Column E
            name: 5,           // Column F
            email: 6,          // Column G
            package: 7,        // Column H
            account_type: 8,   // Column I
            order_type: 9,     // Column J
            revenue_before_discount: 10, // Column K
            discount: 11,      // Column L
            coupon_code: 12,   // Column M
            add_on_amount: 13, // Column N
            add_on_title: 14,  // Column O
            net_revenue: 19,   // Column T
            commission: 20,    // Column U
            order_id: 21       // Column V (in place of response)
          };
  
          if (transactionIdIndex === -1) {
            Logger.log(`Required column 'transaction_id' not found in tab '${tabName}' of sheet '${link}'. Skipping.`);
            return;
          }
  
          // Step 4: Loop Through Local Data and Find Matches
          for (let i = 1; i < localData.length; i++) {
            const customData = localData[i][customDataIndex];
            if (!customData) continue;
  
            const matchedRow = externalData.find(row => row[transactionIdIndex] === customData);
            if (matchedRow) {
              const outputRow = [];
              Object.keys(columnMap).forEach(key => {
                const value = matchedRow[columnMap[key]];
                if (key === "payment_date" && value) {
                  // Format payment date
                  outputRow.push(Utilities.formatDate(new Date(value), Session.getScriptTimeZone(), "dd-MM-yyyy"));
                } else {
                  outputRow.push(value || "");
                }
              });
  
              // Step 5: Write Data to Columns K-Z
              localSheet.getRange(i + 1, 11, 1, outputRow.length).setValues([outputRow]);
            }
          }
        });
      } catch (error) {
        Logger.log(`Error processing sheet link: ${link}, Error: ${error.message}`);
        SpreadsheetApp.getUi().alert(`Error processing sheet link: ${link}`);
      }
    });
  }
  function mainFunctionA(){
    createNuveiTabs();
    // clearOrCreateSheet();
    formatHeaders();
    adjustSalesQIDWorkingTab();
    updateCOAWorkingColumn();
    fetchAndPopulateFinalCOA();
    updateNewDescription();
    updateSalesQIDWorking();
    processPackageAndAccountType();
    generateSalesQIDFinal();
    insertBlankRowsAboveTransactionsOptimized();
    populateHelperColumn();
    shiftTransactionColumnsUp() ;
    filterHelperColumnAndUpdate();
    updateBDOK();
    processAmount();
    processFinalDrCr();
    createFinalQBOTab();
    finetune();
  }
  function mainFunctionB(){
    splitFinalQBOIntoSheets();
    mergeMatchingJournalEntries();
    deleteBlankRowsInSplitTabs();
    uploadSplittedTabsToDrive();
  
  }
  function mainFunctionC(){
    createCustomerTab();
    splitAndUploadCustomerTab();
  }
  //Step-1
  
  function createNuveiTabs() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = ss.getSheetByName("statement");
    
    if (!sourceSheet) {
      SpreadsheetApp.getUi().alert("Source tab 'statement' not found.");
      return;
    }
  
    const data = sourceSheet.getDataRange().getValues();
    if (data.length <= 1) {
      SpreadsheetApp.getUi().alert("No data available in 'statement' tab.");
      return;
    }
  
    const headers = data[0]; // Extract headers
  
    // Helper function to clear or recreate a sheet
    function clearOrCreateSheet(sheetName) {
      let sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        sheet = ss.insertSheet(sheetName);
      } else {
        sheet.clear(); // Clear existing content instead of deleting
      }
      return sheet;
    }
  
    const salesData = [headers];  // Stores sales data
    const refundData = [headers]; // Stores refund & dispute data
  
    // **Step 1: Process data in-memory (Faster)**
    for (let i = 1; i < data.length; i++) {
      if (data[i][6] === 0) {
        salesData.push(data[i]);
      } else {
        refundData.push(data[i]);
      }
    }
  
    // **Step 2: Write Data in Bulk (Faster)**
    const salesSheet = clearOrCreateSheet("Sales");
    const refundSheet = clearOrCreateSheet("Refund+Dispute");
  
    if (salesData.length > 1) salesSheet.getRange(1, 1, salesData.length, salesData[0].length).setValues(salesData);
    if (refundData.length > 1) refundSheet.getRange(1, 1, refundData.length, refundData[0].length).setValues(refundData);
  
    // **Step 3: Optimize Header Formatting**
    formatHeaders(salesSheet, headers.length);
    formatHeaders(refundSheet, headers.length);
  
    SpreadsheetApp.flush(); // Apply all changes at once
    SpreadsheetApp.getUi().alert("Data successfully split into 'Sales' and 'Refund+Dispute' tabs.");
  }
  
  // **Optimized Header Formatting**
  function formatHeaders(sheet, columnCount) {
    if (!sheet || columnCount < 1) return;
  
    let headerRange = sheet.getRange(1, 1, 1, columnCount);
    headerRange.setFontWeight("bold");
  
    // **Step 1: Apply formatting only where required**
    if (columnCount >= 10) {
      let rangeAtoJ = sheet.getRange(1, 1, 1, 10);
      rangeAtoJ.setBackground("#5DB996").setFontColor("black"); // Green header
    }
  
    if (columnCount > 10) {
      let rangeKOnward = sheet.getRange(1, 11, 1, columnCount - 10);
      rangeKOnward.setBackground("yellow").setFontColor("black"); // Yellow header
    }
  }
  
  
  
  
  //Step 2 
  function adjustSalesQIDWorkingTab() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = ss.getSheetByName("Sales");
    if (!sourceSheet) {
      SpreadsheetApp.getUi().alert("Source tab 'Sales' not found.");
      return;
    }
  
    // Delete and recreate the "Sales QID Working" tab
    const tabName = "Sales QID Working";
    let workingSheet = ss.getSheetByName(tabName);
    if (workingSheet) {
      ss.deleteSheet(workingSheet);
    }
    workingSheet = ss.insertSheet(tabName);
  
    // Copy data from "Sales" to "Sales QID Working"
    const data = sourceSheet.getDataRange().getValues();
    workingSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  
    // Insert two new columns between P and Q
    const packageColumnIndex = 16; // Column P
    workingSheet.insertColumnsAfter(packageColumnIndex, 2);
  
    // Set headers for the new columns
    const newHeaders = [["COA Working", "Final COA"]];
    workingSheet.getRange(1, packageColumnIndex + 1, 1, 2).setValues(newHeaders);
  
    // Apply header formatting
    const headerRowRange = workingSheet.getRange(1, 1, 1, workingSheet.getLastColumn());
  
    // A-J: Light green
    if (10 <= headerRowRange.getNumColumns()) {
      workingSheet.getRange(1, 1, 1, 10).setBackground("#5DB996").setFontColor("black");
    }
  
    // K-P: Yellow
    if (packageColumnIndex - 10 > 0) {
      workingSheet.getRange(1, 11, 1, packageColumnIndex - 10).setBackground("yellow").setFontColor("black");
    }
  
    // Q-R: Light yellow
    workingSheet.getRange(1, packageColumnIndex + 1, 1, 2).setBackground("#FFF2C2").setFontColor("black");
  
    // S onward: Yellow
    const lastColumn = workingSheet.getLastColumn();
    if (lastColumn > packageColumnIndex + 2) {
      workingSheet.getRange(1, packageColumnIndex + 3, 1, lastColumn - (packageColumnIndex + 2))
        .setBackground("yellow")
        .setFontColor("black");
    }
  }
  
  //Step -3
  function updateCOAWorkingColumn() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Sales QID Working");
    if (!sheet) {
      SpreadsheetApp.getUi().alert("Sales QID Working tab not found.");
      return;
    }
  
    const data = sheet.getDataRange().getValues(); // Fetch all data
    const packageColIndex = 15; // Column P (zero-based indexing)
    const coaWorkingColIndex = 16; // Column Q (zero-based indexing)
    const accountTypeColIndex = 18; // Column S (zero-based indexing)
    const coaWorkingNewColIndex = 19; // Column R (zero-based indexing)
    const finalCOANewColIndex = 20; // Column S (zero-based indexing)
  
    for (let i = 1; i < data.length; i++) { // Start from row 2 to skip headers
      const packageValue = data[i][packageColIndex]; // Fetch package value
      const accountTypeValue = data[i][accountTypeColIndex]; // Fetch account type value
  
      // Case 1: If "account_type" (S) = 'Swap' and package is not 'NA NAK'
      if (accountTypeValue === 'Swap' && packageValue !== 'NA NAK') {
        data[i][coaWorkingColIndex] = packageValue;
      } 
      // Case 2: If "account_type" (S) = 'Swap Free' and package is not 'NA NAK'
      else if (accountTypeValue === 'Swap Free' && packageValue !== 'NA NAK') {
        data[i][coaWorkingColIndex] = `${packageValue} ${accountTypeValue}`;
      } 
      // Case 3: If "account_type" (S) is blank, paste "Unidentified Transaction (incoming)"
      else if (accountTypeValue === '') {
        data[i][coaWorkingColIndex] = "Unidentified Transaction (incoming)";
      } 
      // Default: Leave "COA Working" blank for unmatched rows
      else {
        data[i][coaWorkingColIndex] = '';
      }
    }
  
    // Write updated data back to the sheet
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    SpreadsheetApp.flush();
    SpreadsheetApp.getUi().alert("COA Working column has been successfully updated!");
  }
  
  //Step-4
  function fetchAndPopulateFinalCOA() {
    try {
      const sheetLink = Browser.inputBox("Enter the external sheet link:");
      if (!sheetLink) {
        SpreadsheetApp.getUi().alert("Sheet link is required. Please try again.");
        return;
      }
  
      const tabName = Browser.inputBox("Enter the external tab name:");
      if (!tabName) {
        SpreadsheetApp.getUi().alert("Tab name is required. Please try again.");
        return;
      }
  
      Logger.log(`Sheet Link: ${sheetLink}`);
      Logger.log(`Tab Name: ${tabName}`);
  
      const externalSpreadsheet = SpreadsheetApp.openByUrl(sheetLink);
      const externalSheet = externalSpreadsheet.getSheetByName(tabName);
      if (!externalSheet) {
        throw new Error(`Tab '${tabName}' not found in the external sheet.`);
      }
  
      const externalData = externalSheet.getRange(2, 2, externalSheet.getLastRow() - 1, 2).getValues(); // Columns B and C
      Logger.log("External Data:");
      Logger.log(externalData);
  
      const externalCOAMapping = {};
      externalData.forEach(row => {
        const phase2COAWorking = row[0];
        const phase2FinalCOA = row[1];
        if (phase2COAWorking) {
          if (!externalCOAMapping[phase2COAWorking]) {
            externalCOAMapping[phase2COAWorking] = [];
          }
          externalCOAMapping[phase2COAWorking].push(phase2FinalCOA);
        }
      });
      Logger.log("External COA Mapping:");
      Logger.log(externalCOAMapping);
  
      const targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sales QID Working");
      if (!targetSheet) {
        throw new Error("Tab 'Sales QID Working' not found in the current spreadsheet.");
      }
  
      const targetData = targetSheet.getDataRange().getValues();
      Logger.log("Target Data (Before):");
      Logger.log(targetData);
  
      const coaWorkingIndex = 16; // Column Q
      const finalCOAIndex = 17; // Column R
  
      for (let i = 1; i < targetData.length; i++) {
        const coaWorkingValue = targetData[i][coaWorkingIndex];
        if (coaWorkingValue && externalCOAMapping[coaWorkingValue]) {
          targetData[i][finalCOAIndex] = externalCOAMapping[coaWorkingValue].join(", ");
        } else {
          targetData[i][finalCOAIndex] = "";
        }
      }
  
      Logger.log("Target Data (After):");
      Logger.log(targetData);
  
      const finalCOARange = targetSheet.getRange(2, finalCOAIndex + 1, targetData.length - 1, 1);
      finalCOARange.setValues(targetData.slice(1).map(row => [row[finalCOAIndex]]));
  
      SpreadsheetApp.flush();
      SpreadsheetApp.getUi().alert("Final COA column has been successfully updated!");
    } catch (error) {
      SpreadsheetApp.getUi().alert(`An error occurred: ${error.message}`);
      Logger.log(`Error: ${error.message}`);
    }
  }
  
  
  //Step -5
  function updateNewDescription() {
    const sheetName = "Sales QID Working";
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
  
    if (!sheet) {
      SpreadsheetApp.getUi().alert(`Sheet '${sheetName}' not found.`);
      return;
    }
  
    // Insert a new column between "name" (Column N) and "email" (Column O)
    const newColumnIndex = 14;
    sheet.insertColumnAfter(newColumnIndex);
    sheet.getRange(1, newColumnIndex + 1).setValue("New Description");
  
    // Get data from the sheet
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      SpreadsheetApp.getUi().alert("No data found to process.");
      return;
    }
  
    const accountTypeIndex = 19;
    const emailIndex = 15;
    const transactionIdIndex = 12;
    const nameIndex = 13;
    const customDataIndex = 9;
    const newDescriptionIndex = newColumnIndex;
  
    // Update data based on conditions
    for (let i = 1; i < data.length; i++) {
      const accountType = data[i][accountTypeIndex];
      const email = data[i][emailIndex] || "";
      const transactionId = data[i][transactionIdIndex] || "";
      const customData = data[i][customDataIndex] || "";
      const isSwap = accountType === 'Swap' || accountType === 'Swap Free';
  
      if (isSwap) {
        data[i][newDescriptionIndex] = `${email}-${transactionId}`;
      } else if (!accountType && customData) {
        data[i][newDescriptionIndex] = `Unidentified Transaction (incoming) - ${customData}`;
      } else {
        data[i][newDescriptionIndex] = "Unidentified Transaction (incoming)";
        // data[i][nameIndex] = "Unidentified - USD";
        // data[i][emailIndex] = "Transaction Unidentified";
      }
      if (!data[i][accountTypeIndex]) { // Check if "account_type" (Column T) is blank
    data[i][nameIndex] = "Unidentified - USD";
    data[i][emailIndex] = "Transaction Unidentified";
  } else {
    // Keep "name" and "email" as they are (do nothing)
  }
  
    }
  
    // Write updated data back to the sheet
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    SpreadsheetApp.flush();
    SpreadsheetApp.getUi().alert("New Description column has been successfully updated!");
  }
  
  //Step -6
  function updateSalesQIDWorking() {
    const sheetName = "Sales QID Working";
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
  
    if (!sheet) {
      SpreadsheetApp.getUi().alert(`Sheet '${sheetName}' not found.`);
      return;
    }
  
    // Insert 7 new columns starting at AD
    sheet.insertColumnsAfter(29, 7);
  
    // Set headers for the new columns
    const headers = ["Total", "Difference", "Nuvei (Dr.)", "Sale Discount (Dr.)", "Add On (Cr.)", "Sales (Cr.)", "Transposed Value"];
    sheet.getRange(1, 30, 1, headers.length).setValues([headers]);
  
    // Get data from the sheet
    const data = sheet.getDataRange().getValues();
    const totalColumnIndex = 29; // Column AD
    const differenceColumnIndex = 30; // Column AE
    const nuveiDrColumnIndex = 31; // Column AF
    const saleDiscountDrColumnIndex = 32; // Column AG
    const addOnCrColumnIndex = 33; // Column AH
    const salesCrColumnIndex = 34; // Column AI
    const transposedValueColumnIndex = 35; // Column AJ
  
    const sourceColumnIndex = 4; // Column E
    const netRevenueColumnIndex = 26; // Column AA
    const discountColumnIndex = 22; // Column W
    const addOnAmountColumnIndex = 24; // Column Y
  
    let transposedValues = []; // Store transposed values for AJ column
  
    for (let i = 1; i < data.length; i++) { // Start from row 2 to skip headers
      const sourceValue = data[i][sourceColumnIndex] || 0; // Default to 0 if blank
      const netRevenueValue = data[i][netRevenueColumnIndex] || 0; // Default to 0 if blank
      const discountValue = data[i][discountColumnIndex] || 0; // Default to 0 if blank
      const addOnAmountValue = data[i][addOnAmountColumnIndex] || 0; // Default to 0 if blank
  
      // Populate columns
      data[i][totalColumnIndex] = sourceValue; // AD = E
      data[i][differenceColumnIndex] = sourceValue - netRevenueValue; // AE = AD - AA
      data[i][nuveiDrColumnIndex] = sourceValue; // AF = AD
      data[i][saleDiscountDrColumnIndex] = discountValue; // AG = W
      data[i][addOnCrColumnIndex] = addOnAmountValue; // AH = Y
      data[i][salesCrColumnIndex] = sourceValue + discountValue - addOnAmountValue; // AI = (AF + AG) - AH
  
      // Add transposed values for AJ column
      transposedValues.push([data[i][nuveiDrColumnIndex]]);
      transposedValues.push([data[i][saleDiscountDrColumnIndex]]);
      transposedValues.push([data[i][addOnCrColumnIndex]]);
      transposedValues.push([data[i][salesCrColumnIndex]]);
    }
  
    // Write updated data back to the sheet (AD to AI)
    sheet.getRange(2, 30, data.length - 1, 6).setValues(data.slice(1).map(row => row.slice(29, 35)));
  
    // Show output for columns AD to AI
    Logger.log("Updated columns AD to AI:", data.slice(1).map(row => row.slice(29, 35)));
  
    // Write transposed values back to AJ column
    const startingRow = 2;
    const totalRows = transposedValues.length;
    sheet.getRange(startingRow, transposedValueColumnIndex + 1, sheet.getLastRow() - 1).clearContent(); // Clear AJ
    sheet.getRange(startingRow, transposedValueColumnIndex + 1, totalRows, 1).setValues(transposedValues);
  
    // Show output for column AJ
    Logger.log("Transposed values in column AJ:", transposedValues);
  
    SpreadsheetApp.flush(); // Ensure all changes are applied
    SpreadsheetApp.getUi().alert("Columns AD to AJ have been successfully updated with calculations and transposed values!");
  }
  //Step -7
  function processPackageAndAccountType() {
    const sheetName = "Sales QID Working";
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
  
    if (!sheet) {
      SpreadsheetApp.getUi().alert(`Sheet '${sheetName}' not found.`);
      return;
    }
  
    // Get data from the sheet
    const data = sheet.getDataRange().getValues();
    const packageColIndex = 16; // Column Q (zero-based indexing)
    const nameColIndex = 13; // Column N (zero-based indexing)
    const newDescriptionColIndex = 14; // Column O (zero-based indexing)
    const emailColIndex = 15; // Column P (zero-based indexing)
    const coaWorkingColIndex = 17; // Column R (zero-based indexing)
    const finalCOAColIndex = 18; // Column S (zero-based indexing)
    const accountTypeColIndex = 19; // Column T (zero-based indexing)
  
    for (let i = 1; i < data.length; i++) { // Start from row 2 to skip headers
      const packageValue = data[i][packageColIndex];
      const accountTypeValue = data[i][accountTypeColIndex];
  
  // Handle case for 'Swap'
   if (packageValue === 'NA NAK' && accountTypeValue === 'Swap') {
      data[i][nameColIndex] = ""; // Blank name
      data[i][newDescriptionColIndex] = ""; // Blank New Description
      data[i][emailColIndex] = ""; // Blank email
      data[i][coaWorkingColIndex] = "Paid Competition"; // COA Working
      data[i][finalCOAColIndex] = "Trading Income (FundedNext):Paid Competition"; // Final COA
  }
  
  // Handle case for 'Swap Free'
   else if (packageValue === 'NA NAK' && accountTypeValue === 'Swap Free') {
      data[i][nameColIndex] = ""; // Blank name
      data[i][newDescriptionColIndex] = ""; // Blank New Description
      data[i][emailColIndex] = ""; // Blank email
      data[i][coaWorkingColIndex] = "Paid Competition Swap Free"; // COA Working
      data[i][finalCOAColIndex] = "Trading Income (FundedNext):Paid Competition Swap Free"; // Final COA
  }
  
    }
  
    // Write updated data back to the sheet
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  
    SpreadsheetApp.getUi().alert("Processing complete for 'NA NAK' with account types 'Swap' and 'Swap Free'.");
  }
  
  //Step -8
  
  function generateSalesQIDFinal() {
    const sourceSheetName = "Sales QID Working";
    const targetSheetName = "Sales QID Final";
    const ss = SpreadsheetApp.getActiveSpreadsheet();
  
    // Get the source sheet
    const sourceSheet = ss.getSheetByName(sourceSheetName);
    if (!sourceSheet) {
      SpreadsheetApp.getUi().alert(`Source sheet '${sourceSheetName}' not found.`);
      return;
    }
  
    // Check if target sheet exists, if not, create it
    let targetSheet = ss.getSheetByName(targetSheetName);
    if (!targetSheet) {
      targetSheet = ss.insertSheet(targetSheetName);
    } else {
      targetSheet.clear(); // Clear existing data
    }
  
    // Set headers in the target sheet
    const headers = [
      "*JournalNo", "", "*JournalDate", "", "*AccountName", "", "*Debits", "*Credits", 
      "Description", "Name", "", "Currency", "Location", "Class", "", "Memo"
    ];
    targetSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    targetSheet.getRange(1, 1, 1, headers.length).setBackground("yellow"); // Highlight headers
  
    // Get data from the source sheet
    const sourceData = sourceSheet.getDataRange().getValues();
    if (sourceData.length <= 1) {
      SpreadsheetApp.getUi().alert("No data available in the source sheet.");
      return;
    }
  
    // Filter rows needed from the source sheet
    const filteredData = sourceData.filter((row, index) => {
      if (index === 0) return false; // Skip headers
      return row[18] && row[3]; // Ensure 'Final COA' (S) and 'Date' (D) are not blank
    });
  
    // Prompt for the first journal number
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt("Enter the first journal number (e.g., NUV-73598):");
    if (response.getSelectedButton() !== ui.Button.OK || !response.getResponseText()) {
      ui.alert("Invalid or no journal number provided. Aborting process.");
      return;
    }
    let currentJournalNo = response.getResponseText().trim();
  
    // Validate journal number format
    const validJournalNo = /^[a-zA-Z0-9]+-\d+$/; // Example: "NUV-73598"
    if (!validJournalNo.test(currentJournalNo)) {
      ui.alert("Invalid journal number format. Please use a format like 'NUV-73598'.");
      return;
    }
  
    // Extract prefix and last number
    let journalPrefix = currentJournalNo.replace(/\d+$/, ""); // Extract prefix
    let journalNumber = parseInt(currentJournalNo.match(/\d+$/)[0], 10); // Extract number
  
    const targetData = filteredData.map((row) => {
      const journalDate = row[3]; // 'Date' (Column D in source sheet)
      const accountName = row[18]; // 'Final COA' (Column S in source sheet)
      const name = row[13]; // 'Name' (Column N in source sheet)
      const memo = row[5]; // 'Card Holder Email' (Column F in source sheet)
      const description = row[14]; // 'New Description' (Column O in source sheet)
      const classValue = row[20]; // 'order_type' (Column U in source sheet)
  
      // Assign journal number and increment
      const journalNo = `${journalPrefix}${journalNumber}`;
      journalNumber++; // Increment journal number for the next row
  
      return [
        journalNo, "", journalDate, "", accountName, "", "", "", 
        description, name, "", "USD", "", classValue, "", memo
      ];
    });
  
    // Write data to the target sheet
    targetSheet.getRange(2, 1, targetData.length, headers.length).setValues(targetData);
  
    SpreadsheetApp.flush(); // Ensure all changes are applied
    SpreadsheetApp.getUi().alert("'Sales QID Final' tab has been successfully created and populated.");
  }
  
  
  //Step -9
  
  function insertBlankRowsAboveTransactionsOptimized() {
    const sheetName = "Sales QID Final";
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
  
    if (!sheet) {
      SpreadsheetApp.getUi().alert(`Sheet '${sheetName}' not found.`);
      return;
    }
  
    // Get all data from the sheet
    const data = sheet.getDataRange().getValues();
    const numColumns = data[0].length; // Get total columns
    const newData = [data[0]]; // Keep headers
  
    let lastJournalNo = ""; // Store last assigned journal number
  
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] !== "") { 
        lastJournalNo = data[i][0]; // Capture the last valid journal number
      }
  
      // **Insert blank rows only if the previous row was NOT already blank**
      if (newData.length === 1 || newData[newData.length - 1].some(cell => cell !== "")) {
        newData.push(new Array(numColumns).fill("")); // Blank Row 1
        newData.push(new Array(numColumns).fill("")); // Blank Row 2
        newData.push(new Array(numColumns).fill("")); // Blank Row 3
      }
  
      newData.push(data[i]); // Add the transaction row
    }
  
    // **Optimized Update: Only Write to Affected Range**
    const numRows = newData.length;
    sheet.getRange(1, 1, sheet.getLastRow(), numColumns).clearContent(); // Clear only required area
    sheet.getRange(1, 1, numRows, numColumns).setValues(newData); // Write back formatted data
  
    SpreadsheetApp.flush();
    SpreadsheetApp.getUi().alert("Blank rows added above each transaction successfully.");
  }
  
  //Step-10
  
  function populateHelperColumn() {
    const sheetName = "Sales QID Final"; 
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
  
    if (!sheet) {
      SpreadsheetApp.getUi().alert(`Sheet '${sheetName}' not found.`);
      return;
    }
  
    // **Step 1: Set header for column F**
    sheet.getRange(1, 6).setValue("Helper");
  
    // **Step 2: Get all data from *JournalNo (Column A)**
    const data = sheet.getDataRange().getValues(); // Fetch all rows
    const helperColumn = [];
    let lastTransactionIndex = -1; // Track last transaction row
  
    // **Step 3: Identify the last transaction row in Column A**
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][0]) { // If *JournalNo is NOT empty
        lastTransactionIndex = i;
        break;
      }
    }
  
    if (lastTransactionIndex === -1) {
      SpreadsheetApp.getUi().alert("No transactions found in *JournalNo (Column A).");
      return;
    }
  
    // **Step 4: Clear Column F before regenerating**
    sheet.getRange(2, 6, data.length - 1).clearContent(); 
  
    // **Step 5: Generate Helper values in cyclic order [1,2,3,0]**
    let cycleIndex = 1; // Start with 1
    for (let i = 1; i <= lastTransactionIndex; i++) {
      if (data[i][0]) { 
        helperColumn.push([0]); // Assign '0' to rows with a transaction in Column A
        cycleIndex = 1; // Reset cycle after transaction
      } else {
        helperColumn.push([cycleIndex]); // Assign cyclic values
        cycleIndex = (cycleIndex % 3) + 1; // Rotate through 1,2,3
      }
    }
  
    // **Step 6: Write updated values into column F**
    sheet.getRange(2, 6, helperColumn.length, 1).setValues(helperColumn);
  
    SpreadsheetApp.flush();
    SpreadsheetApp.getUi().alert("Helper column has been successfully updated with the [1,2,3,0] cycle!");
  }
  
  //step-11
  
  function shiftTransactionColumnsUp() {
    const sheetName = "Sales QID Final";
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
  
    if (!sheet) {
      SpreadsheetApp.getUi().alert(`Sheet '${sheetName}' not found.`);
      return;
    }
  
    const numRowsToShift = 3; // Number of rows to move up
    const totalRows = sheet.getLastRow(); // Get total row count
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
    // Identify column indexes dynamically based on header names
    const columnsToShift = {
      journalNo: headers.indexOf("*JournalNo") + 1,
      journalDate: headers.indexOf("*JournalDate") + 1,
      name: headers.indexOf("Name") + 1,
      currency: headers.indexOf("Currency") + 1,
      class: headers.indexOf("Class") + 1,
      memo: headers.indexOf("Memo") + 1
    };
  
    // Ensure all necessary columns are found
    if (Object.values(columnsToShift).includes(0)) {
      SpreadsheetApp.getUi().alert("One or more required columns are missing. Please check column headers.");
      return;
    }
  
    // Get data for the selected columns
    let dataToMove = {};
    for (const [key, col] of Object.entries(columnsToShift)) {
      dataToMove[key] = sheet.getRange(2, col, totalRows - 1).getValues(); // Exclude header row
    }
  
    // Clear original locations (Cut operation)
    for (const col of Object.values(columnsToShift)) {
      sheet.getRange(2, col, totalRows - 1).clearContent();
    }
  
    // Paste values 3 rows up
    for (const [key, col] of Object.entries(columnsToShift)) {
      const targetRange = sheet.getRange(2, col, totalRows - numRowsToShift - 1);
      targetRange.setValues(dataToMove[key].slice(numRowsToShift)); // Paste shifted values
    }
  
    SpreadsheetApp.flush();
    SpreadsheetApp.getUi().alert("Transaction columns have been shifted up successfully by 3 rows!");
  }
  
  //Step-12
  function filterHelperColumnAndUpdate() {
    const sheetName = "Sales QID Final"; // Target sheet name
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
  
    if (!sheet) {
      SpreadsheetApp.getUi().alert(`Sheet '${sheetName}' not found.`);
      return;
    }
  
    // Get all data from the sheet
    const data = sheet.getDataRange().getValues();
    
    // Define column indexes
    const helperColIndex = 5; // Column F ("Helper") - Zero-based
    const accountNameColIndex = 4; // Column E ("*AccountName") - Zero-based
    const descriptionColIndex = 8; // Column I ("Description") - Zero-based
  
    // Loop through each row, starting from row 2 (excluding headers)
    for (let i = 1; i < data.length; i++) {
      const helperValue = data[i][helperColIndex]; // Get Helper column value
  
      if (helperValue === 1) {
        // Rule 1: If Helper is 1
        data[i][accountNameColIndex] = "Nuvei - Safe Charge (USD)";
        data[i][descriptionColIndex] = "Nuvei - Safe Charge (USD) - Revenue";
      } else if (helperValue === 2) {
        // Rule 2: If Helper is 2
        data[i][accountNameColIndex] = "Sales Discount";
        data[i][descriptionColIndex] = "Sales Discount";
      } else if (helperValue === 3) {
        // Rule 3: If Helper is 3
        data[i][accountNameColIndex] = "Trading Income (FundedNext):Addon";
        data[i][descriptionColIndex] = "Addon Income";
      }
    }
  
    // Write updated data back to the sheet
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    SpreadsheetApp.flush();
  
    // Alert user that the process is complete
    SpreadsheetApp.getUi().alert("Filtering and updates based on the 'Helper' column have been successfully applied!");
  }
  
  
  // step -13
  function updateBDOK() {
     const ss = SpreadsheetApp.getActiveSpreadsheet();
  
    const sourceSheetName = "Sales QID Working";
    const targetSheetName = "Sales QID Final";
  
    const sourceSheet = ss.getSheetByName(sourceSheetName);
    const targetSheet = ss.getSheetByName(targetSheetName);
  
    if (!sourceSheet || !targetSheet) {
      SpreadsheetApp.getUi().alert("Source or target sheet not found.");
      return;
    }
  
    const sourceData = sourceSheet.getDataRange().getValues();
    const targetData = targetSheet.getDataRange().getValues();
  
    const orderTypeIndex = 20; // 'order_type' is column U (zero-based index)
    const classColumnIndex = 13; // 'Class' is column N in 'Sales QID Final' (zero-based index)
  
    // Validate data length
    if (sourceData.length <= 1 || targetData.length <= 1) {
      SpreadsheetApp.getUi().alert("No data found in source or target sheet.");
      return;
    }
  
    // Write updated data back to 'Sales QID Final'
    targetSheet.getRange(1, 1, targetData.length, targetData[0].length).setValues(targetData);
  
    // Apply formulas to columns B, D, K, O
    const lastRow = targetSheet.getLastRow();
    targetSheet.getRange(2, 2, lastRow - 1, 1).setFormula("=IF(A2=\"\",B1,A2)"); // Column B
    targetSheet.getRange(2, 4, lastRow - 1, 1).setFormula("=IF(C2=\"\",D1,C2)"); // Column D
    targetSheet.getRange(2, 11, lastRow - 1, 1).setFormula("=IF(J2=\"\",K1,J2)"); // Column K
    targetSheet.getRange(2, 15, lastRow - 1, 1).setFormula("=IF(N2=\"\",O1,N2)"); // Column O
  
    SpreadsheetApp.flush();
    SpreadsheetApp.getUi().alert("Data and formulas updated successfully in 'Sales QID Final'.");
  }
  
  
  //Step -13
  function processAmount(){
    const sourceSheetName = "Sales QID Working";
    const targetSheetName = "Sales QID Final";
    const ss = SpreadsheetApp.getActiveSpreadsheet();
  
    // Get source and target sheets
    const sourceSheet = ss.getSheetByName(sourceSheetName);
    const targetSheet = ss.getSheetByName(targetSheetName);
  
    if (!sourceSheet || !targetSheet) {
      SpreadsheetApp.getUi().alert(`One or both sheets ('${sourceSheetName}' or '${targetSheetName}') not found.`);
      return;
    }
  
    // Determine last row in source sheet based on "Transposed Value" (Column AJ = Index 36)
    const lastRow = sourceSheet.getLastRow();
    if (lastRow < 2) {
      SpreadsheetApp.getUi().alert("No data found in the source sheet.");
      return;
    }
  
    // Copy "Transposed Value" (AJ) data from 'Sales QID Working' (Column 36)
    const transposedValues = sourceSheet.getRange(2, 36, lastRow - 1, 1).getValues();
  
    // Paste into "*Debits" (G) in 'Sales QID Final' (Column 7)
    targetSheet.getRange(2, 7, transposedValues.length, 1).setValues(transposedValues);
  
    SpreadsheetApp.flush(); // Apply changes
    SpreadsheetApp.getUi().alert("Processing complete: Transposed Values copied to *Debits.");
  }
  
  //Step-14
  
  function processFinalDrCr() {
    const sheetName = "Sales QID Final";
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
  
    if (!sheet) {
      SpreadsheetApp.getUi().alert(`Sheet '${sheetName}' not found.`);
      return;
    }
  
    // Get all data
    const dataRange = sheet.getDataRange();
    const data = dataRange.getValues();
  
    const helperColIndex = 5; // Column F (Helper)
    const debitsColIndex = 6; // Column G (*Debits)
    const creditsColIndex = 7; // Column H (*Credits)
  
    // Process each row (starting from row 2)
    for (let i = 1; i < data.length; i++) {
      const helperValue = data[i][helperColIndex];
  
      if (helperValue === 0 || helperValue === 3) {
        // Move Debits (G) to Credits (H)
        data[i][creditsColIndex] = data[i][debitsColIndex];
        data[i][debitsColIndex] = ""; // Clear Debits (G)
      }
    }
  
    // Write updated data back to "Sales QID Final"
    dataRange.setValues(data);
  
    SpreadsheetApp.flush(); // Apply all changes
    SpreadsheetApp.getUi().alert("Final Dr/Cr processing completed successfully!");
  }
  
  //Step -15
  
  function createFinalQBOTab() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheetName = "Sales QID Final";
    const targetSheetName = "Final QBO";
  
    // Get source sheet
    const sourceSheet = ss.getSheetByName(sourceSheetName);
    if (!sourceSheet) {
      SpreadsheetApp.getUi().alert(`Source sheet '${sourceSheetName}' not found.`);
      return;
    }
  
    // Delete and recreate "Final QBO" sheet if it exists
    let targetSheet = ss.getSheetByName(targetSheetName);
    if (targetSheet) {
      ss.deleteSheet(targetSheet);
    }
    targetSheet = ss.insertSheet(targetSheetName);
  
    // Copy all data from "Sales QID Final" to "Final QBO"
    const dataRange = sourceSheet.getDataRange();
    const data = dataRange.getValues();
    targetSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  
    // Apply formatting from source to target
    const sourceFormats = dataRange.getFontWeights();
    targetSheet.getRange(1, 1, data.length, data[0].length).setFontWeights(sourceFormats);
  
    // Identify columns
    const headers = data[0];
    const debitColumnIndex = headers.indexOf("*Debits");
    const creditColumnIndex = headers.indexOf("*Credits");
  
    if (debitColumnIndex === -1 || creditColumnIndex === -1) {
      SpreadsheetApp.getUi().alert("Required columns (*Debits or *Credits) not found.");
      return;
    }
  
    // Function to remove rows where the column has '0'
    function removeRowsByColumn(sheet, columnIndex) {
      const range = sheet.getDataRange();
      let values = range.getValues();
      let filteredData = values.filter((row, index) => index === 0 || row[columnIndex] !== 0); // Keep header row and non-zero rows
      sheet.clear();
      sheet.getRange(1, 1, filteredData.length, filteredData[0].length).setValues(filteredData);
    }
  
    // Step 1: Delete rows where "*Debits*" (G) = 0
    removeRowsByColumn(targetSheet, debitColumnIndex);
  
    // Step 2: Delete rows where "*Credits*" (H) = 0
    removeRowsByColumn(targetSheet, creditColumnIndex);
    SpreadsheetApp.flush();
    SpreadsheetApp.getUi().alert("Final QBO tab has been successfully created with filtered and adjusted data!");
    
    };
  
  //Step-16
  
  function finetune() {
    const sheetName = "Final QBO";
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
  
    if (!sheet) {
      SpreadsheetApp.getUi().alert(`Sheet '${sheetName}' not found.`);
      return;
    }
  
    // Get headers from row 1
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
    // Find column indexes dynamically
    const colA = headers.indexOf("*JournalNo") + 1;  // Column A
    const colB = colA + 1; // Column B (where A will be copied)
    const colC = headers.indexOf("*JournalDate") + 1; // Column C
    const colD = colC + 1; // Column D (where C will be copied)
    const colJ = headers.indexOf("Name") + 1; // Column J
    const colK = colJ + 1; // Column K (where J will be copied)
    const colN = headers.indexOf("Class") + 1; // Column N
    const colO = colN + 1; // Column O (where N will be copied)
  
    // Ensure all necessary columns exist
    if ([colA, colB, colC, colD, colJ, colK, colN, colO].includes(0)) {
      SpreadsheetApp.getUi().alert("One or more required columns are missing. Please check the sheet.");
      return;
    }
  
    // Copy headers
    sheet.getRange(1, colB).setValue(headers[colA - 1]); // Copy A to B
    sheet.getRange(1, colD).setValue(headers[colC - 1]); // Copy C to D
    sheet.getRange(1, colK).setValue(headers[colJ - 1]); // Copy J to K
    sheet.getRange(1, colO).setValue(headers[colN - 1]); // Copy N to O
  
    SpreadsheetApp.flush();
  
    // Confirm before deleting columns
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      "Confirm Deletion",
      "Are you sure you want to delete columns A, D, F, J, and N?",
      ui.ButtonSet.YES_NO
    );
  
    if (response === ui.Button.YES) {
      // Delete columns in reverse order to avoid index shifts
      const columnsToDelete = [colA, colD, headers.indexOf("Helper") + 1, colJ, colN].filter(index => index > 0);
      columnsToDelete.sort((a, b) => b - a); // Sort in descending order
  
      columnsToDelete.forEach(index => sheet.deleteColumn(index));
  
      SpreadsheetApp.flush();
      SpreadsheetApp.getUi().alert("Columns deleted successfully.");
    } else {
      SpreadsheetApp.getUi().alert("Column deletion cancelled.");
    }
  }
  
  //Step- 17
  function splitFinalQBOIntoSheets() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = ss.getSheetByName("Final QBO");
  
    if (!sourceSheet) {
      SpreadsheetApp.getUi().alert(" Sheet 'Final QBO' not found.");
      return;
    }
  
    const data = sourceSheet.getDataRange().getValues();
    if (data.length <= 1) {
      SpreadsheetApp.getUi().alert("No data found in 'Final QBO'.");
      return;
    }
  
    const headerRow = data[0]; // Store headers
    const journalNoIndex = 0; // Column A (*JournalNo) (0-based index)
    const maxRowsPerSheet = 990; // Each sheet will contain 990 rows (excluding header)
    let validRows = [];
  
    // **Step 1: Collect non-null rows from *JournalNo (Column A)**
    for (let i = 1; i < data.length; i++) { // Skip header row
      if (data[i][journalNoIndex]) {
        validRows.push(data[i]); // Only add rows with non-null *JournalNo
      }
    }
  
    // **Step 2: Calculate number of sheets needed**
    const totalRows = validRows.length;
    if (totalRows === 0) {
      SpreadsheetApp.getUi().alert("No valid data to split.");
      return;
    }
    
    const numSheets = Math.ceil(totalRows / maxRowsPerSheet);
  
    // **Step 3:Delete all previous Split_X sheets in one go (Faster)**
    const existingSheets = ss.getSheets();
    const sheetsToDelete = existingSheets.filter(sheet => /^Split_\d+$/.test(sheet.getName()));
    sheetsToDelete.forEach(sheet => ss.deleteSheet(sheet));
  
    // **Step 4: Create all sheets first**
    let newSheets = [];
    for (let i = 0; i < numSheets; i++) {
      let newSheet = ss.insertSheet(`Split_${i + 1}`);
      newSheets.push(newSheet);
    }
  
    // **Step 5: Write data in bulk**
    for (let i = 0; i < numSheets; i++) {
      let startRow = i * maxRowsPerSheet;
      let endRow = Math.min(startRow + maxRowsPerSheet, totalRows);
      let chunk = validRows.slice(startRow, endRow);
  
      let sheet = newSheets[i]; // Get the corresponding sheet
      sheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]); // Write headers
      sheet.getRange(2, 1, chunk.length, headerRow.length).setValues(chunk); // Write data
    }
  
    SpreadsheetApp.flush();
    SpreadsheetApp.getUi().alert(`Splitting Complete! Created ${numSheets} sheets with up to 990 rows each.`);
  }
  
  
  //step -18
  function mergeMatchingJournalEntries() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheetIndex = 1; // Start from "Split_2"
  
    while (true) {
      let prevSheetName = `Split_${sheetIndex}`;
      let currentSheetName = `Split_${sheetIndex + 1}`;
  
      let prevSheet = ss.getSheetByName(prevSheetName);
      let currentSheet = ss.getSheetByName(currentSheetName);
  
      if (!prevSheet || !currentSheet) break; // Stop if no more sheets exist
  
      // Get the *JournalNo value from A991 in the previous sheet
      let lastJournalNo = prevSheet.getRange(991, 1).getValue().toString().trim();
  
      if (lastJournalNo) {
        let currentData = currentSheet.getDataRange().getValues();
        let rowsToMove = [];
  
        // Identify rows with matching *JournalNo
        for (let i = 1; i < currentData.length; i++) {
          if (currentData[i][0].toString().trim() === lastJournalNo) {
            rowsToMove.push(currentData[i]);
            currentSheet.getRange(i + 1, 1, 1, currentSheet.getLastColumn()).clearContent(); // Clear the row
          }
        }
  
        if (rowsToMove.length > 0) {
          let prevLastRow = prevSheet.getLastRow() + 1;
          prevSheet.getRange(prevLastRow, 1, rowsToMove.length, rowsToMove[0].length).setValues(rowsToMove);
        }
      }
  
      sheetIndex++; // Move to the next sheet
    }
  
    SpreadsheetApp.getUi().alert("Journal entries merged successfully!");
  }
  
  //Step 19
  function deleteBlankRowsInSplitTabs() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets();
  
    // **Step 1: Filter only "Split_X" tabs**
    var splitSheets = sheets.filter(sheet => /^Split_\d+$/.test(sheet.getName()));
  
    if (splitSheets.length === 0) {
      SpreadsheetApp.getUi().alert("No 'Split_X' tabs found.");
      return;
    }
  
    var totalDeleted = 0;
  
    // **Step 2: Process each "Split_X" sheet**
    splitSheets.forEach(sheet => {
      var data = sheet.getDataRange().getValues(); // Get all data
      var journalNoColumnIndex = 0; // Column A (0-based index)
      var rowsToDelete = [];
  
      // **Step 3: Loop through all rows, collect indexes of blank "*JournalNo*" rows**
      for (var i = data.length - 1; i > 0; i--) { // Start from bottom to prevent shifting issues
        if (!data[i][journalNoColumnIndex]) { // If "*JournalNo*" is blank
          rowsToDelete.push(i + 1); // Store row index (1-based index)
        }
      }
  
      // **Step 4: Delete rows in reverse order to avoid shifting issues**
      if (rowsToDelete.length > 0) {
        rowsToDelete.forEach(row => sheet.deleteRow(row));
        totalDeleted += rowsToDelete.length;
      }
    });
  
    // **Step 5: Alert user about the result**
    if (totalDeleted > 0) {
      SpreadsheetApp.getUi().alert(totalDeleted + " blank row(s) deleted from 'Split_X' tabs.");
    } else {
      SpreadsheetApp.getUi().alert("No blank rows found in 'Split_X' tabs.");
    }
  
    SpreadsheetApp.flush(); // Ensure all changes are applied
  }
  //Step 20
  function uploadSplittedTabsToDrive() {
    var ui = SpreadsheetApp.getUi();
  
    // Step 1: Prompt user for Google Drive folder URL
    var response = ui.prompt("Enter the Google Drive folder URL where you want to save the CSV files:");
  
    // If user cancels or doesn't enter anything, stop execution
    if (response.getSelectedButton() !== ui.Button.OK || !response.getResponseText()) {
      ui.alert("No folder URL provided. Process aborted.");
      return;
    }
  
    var folderUrl = response.getResponseText().trim();
  
    // Extract Folder ID from URL
    var folderIdMatch = folderUrl.match(/[-\w]{25,}/);
    if (!folderIdMatch) {
      ui.alert("Invalid Google Drive folder URL. Please enter a valid URL.");
      return;
    }
    var folderId = folderIdMatch[0];
  
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets();
  
    // Step 2: Filter only "Split_X" tabs
    var splitSheets = sheets.filter(sheet => /^Split_\d+$/.test(sheet.getName()));
  
    if (splitSheets.length === 0) {
      ui.alert("No 'Split_X' sheets found to upload.");
      return;
    }
  
    // Step 3: Convert each Split_X sheet into a CSV file and upload using Drive API
    splitSheets.forEach(sheet => {
      var data = sheet.getDataRange().getValues();
      if (data.length <= 1) return; // Skip if there's only the header row
  
      var csvContent = data.map(row => row.join(",")).join("\n");
      var fileName = sheet.getName() + ".csv";
      var blob = Utilities.newBlob(csvContent, "text/csv", fileName);
  
      // Use Drive API instead of DriveApp
      var fileMetadata = {
        name: fileName,
        mimeType: "text/csv",
        parents: [folderId] // Ensures the file is saved in the specified folder
      };
  
      var uploadedFile = Drive.Files.create(fileMetadata, blob, {
        supportsAllDrives: true
      });
  
      Logger.log("Uploaded: " + uploadedFile.id);
    });
  
    // Step 4: Alert user with folder link
    ui.alert(`All 'Split_X' tabs have been successfully uploaded to:\n${folderUrl}`);
  }
  
  
  //step 21
  function processCustomerData() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var salesSheet = ss.getSheetByName("Sales QID Final");
    var customerSheet = ss.getSheetByName("Customer");
  
    // Step 1: Create "Customer" tab if it doesn't exist
    if (!customerSheet) {
      customerSheet = ss.insertSheet("Customer");
    } else {
      customerSheet.clear(); // Clear previous data
    }
  
    // Step 2: Get data from "Sales QID Final"
    var data = salesSheet.getDataRange().getValues();
    var headers = data[0]; // Header row
    var helperColIndex = headers.indexOf("Helper");
    var descriptionColIndex = headers.indexOf("Description");
    var memoColIndex = headers.indexOf("Memo");
  
    if (helperColIndex === -1 || descriptionColIndex === -1 || memoColIndex === -1) {
      SpreadsheetApp.getUi().alert("One or more required columns are missing in 'Sales QID Final' tab.");
      return;
    }
  
    var filteredData = [];
    filteredData.push(["Description", "Memo"]); // Headers for Customer tab
  
    // Step 3: Filter data where "Helper" (F) column = 1
    for (var i = 1; i < data.length; i++) {
      if (data[i][helperColIndex] === 1) {
        filteredData.push([data[i][descriptionColIndex], data[i][memoColIndex]]);
      }
    }
  
    // Step 4: Paste filtered data into "Customer" tab
    customerSheet.getRange(1, 1, filteredData.length, 2).setValues(filteredData);
  
    if (filteredData.length === 1) {
      SpreadsheetApp.getUi().alert("No data found where 'Helper' is 1.");
      return;
    }
  
    // Step 5: Split "Customer" tab into sheets with 990 rows each
    splitAndUploadCustomerTab(ss, customerSheet, filteredData);
  }
  
  function splitAndUploadCustomerTab(ss, sourceSheet, data) {
    var headerRow = data[0]; // Store headers
    var maxRowsPerSheet = 990; // Split size
    var validRows = data.slice(1); // Remove header from data array
  
    // Step 6: Calculate number of split sheets
    var totalRows = validRows.length;
    if (totalRows === 0) {
      SpreadsheetApp.getUi().alert("No data to split.");
      return;
    }
    
    var numSheets = Math.ceil(totalRows / maxRowsPerSheet);
  
    // Step 7: Delete existing "Customer_Split_X" tabs, but only relevant ones
    ss.getSheets().forEach(sheet => {
      if (/^Customer_Split_\d+$/.test(sheet.getName())) { // Strict match for split sheets
        ss.deleteSheet(sheet);
      }
    });
  
    // Step 8: Create and populate split sheets (only if there is data)
    for (var i = 0; i < numSheets; i++) {
      let startRow = i * maxRowsPerSheet;
      let endRow = Math.min(startRow + maxRowsPerSheet, totalRows);
  
      let chunk = validRows.slice(startRow, endRow);
      if (chunk.length === 0) continue; // Avoid creating unnecessary blank sheets
  
      let newSheetName = `Customer_Split_${i + 1}`;
      let newSheet = ss.getSheetByName(newSheetName) || ss.insertSheet(newSheetName);
  
      // Clear previous data
      newSheet.clear();
  
      // Write headers
      newSheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);
  
      // Write split data (990 rows per sheet)
      newSheet.getRange(2, 1, chunk.length, headerRow.length).setValues(chunk);
    }
  
    // Step 9: Ask user for the Google Drive folder URL
    var ui = SpreadsheetApp.getUi();
    var response = ui.prompt("Enter Google Drive folder URL where CSV files will be uploaded:");
    var folderUrl = response.getResponseText().trim();
  
    if (!folderUrl) {
      ui.alert("No folder URL provided. Process aborted.");
      return;
    }
  
    try {
      var folderId = folderUrl.match(/[-\w]{25,}/)[0]; // Extract folder ID from URL
      var folder = DriveApp.getFolderById(folderId);
    } catch (e) {
      ui.alert("Invalid Google Drive folder URL. Please provide a correct URL.");
      return;
    }
  
    // Step 10: Convert each split sheet into CSV and upload (only non-empty sheets)
    var splitSheets = ss.getSheets().filter(sheet => /^Customer_Split_\d+$/.test(sheet.getName()));
  
    splitSheets.forEach(sheet => {
      var values = sheet.getDataRange().getValues();
      if (values.length <= 1) return; // Skip if only header row is present
  
      var csvData = values.map(row => row.map(value => `"${value}"`).join(",")).join("\n");
      var csvBlob = Utilities.newBlob(csvData, "text/csv", sheet.getName() + ".csv");
  
      folder.createFile(csvBlob);
    });
  
    SpreadsheetApp.getUi().alert(`✅ Splitting & Upload Complete! Created ${numSheets} CSV files in the provided Google Drive folder.`);
  }
  
  
  
  //Last step
  function processCustomerData() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var salesSheet = ss.getSheetByName("Sales QID Final");
    var customerSheet = ss.getSheetByName("Customer");
  
    // Step 1: Create "Customer" tab if it doesn't exist
    if (!customerSheet) {
      customerSheet = ss.insertSheet("Customer");
    } else {
      customerSheet.clear(); // Clear previous data
    }
  
    // Step 2: Get data from "Sales QID Final"
    var data = salesSheet.getDataRange().getValues();
    var headers = data[0]; // Header row
    var helperColIndex = headers.indexOf("Helper");
    var descriptionColIndex = headers.indexOf("Description");
    var memoColIndex = headers.indexOf("Memo");
  
    if (helperColIndex === -1 || descriptionColIndex === -1 || memoColIndex === -1) {
      SpreadsheetApp.getUi().alert("One or more required columns are missing in 'Sales QID Final' tab.");
      return;
    }
  
    var filteredData = [];
    filteredData.push(["Description", "Memo"]); // Headers for Customer tab
  
    // Step 3: Filter data where "Helper" (F) column = 1
    for (var i = 1; i < data.length; i++) {
      if (data[i][helperColIndex] === 1) {
        filteredData.push([data[i][descriptionColIndex], data[i][memoColIndex]]);
      }
    }
  
    // Step 4: Paste filtered data into "Customer" tab
    customerSheet.getRange(1, 1, filteredData.length, 2).setValues(filteredData);
  
    if (filteredData.length === 1) {
      SpreadsheetApp.getUi().alert("No data found where 'Helper' is 1.");
      return;
    }
  
    // Step 5: Split "Customer" tab into sheets with 990 rows each
    splitAndUploadCustomerTab(ss, customerSheet, filteredData);
  }
  
  function splitAndUploadCustomerTab(ss, sourceSheet, data) {
    var headerRow = data[0]; // Store headers
    var maxRowsPerSheet = 990; // Split size
    var validRows = data.slice(1); // Remove header from data array
  
    // Step 6: Calculate number of split sheets
    var totalRows = validRows.length;
    var numSheets = Math.ceil(totalRows / maxRowsPerSheet);
  
    // Step 7: Delete existing "Customer_Split_X" tabs
    ss.getSheets().forEach(sheet => {
      if (sheet.getName().startsWith("Customer_Split_")) {
        ss.deleteSheet(sheet);
      }
    });
  
    // Step 8: Create and populate split sheets
    for (var i = 0; i < numSheets; i++) {
      let startRow = i * maxRowsPerSheet;
      let endRow = Math.min(startRow + maxRowsPerSheet, totalRows);
  
      let newSheetName = `Customer_Split_${i + 1}`;
      let newSheet = ss.getSheetByName(newSheetName) || ss.insertSheet(newSheetName);
  
      // Clear previous data
      newSheet.clear();
  
      // Write headers
      newSheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);
  
      // Write split data (990 rows per sheet)
      let chunk = validRows.slice(startRow, endRow);
      newSheet.getRange(2, 1, chunk.length, headerRow.length).setValues(chunk);
    }
  
    // Step 9: Ask user for the Google Drive folder URL
    var ui = SpreadsheetApp.getUi();
    var response = ui.prompt("Enter Google Drive folder URL where CSV files will be uploaded:");
    var folderUrl = response.getResponseText().trim();
  
    if (!folderUrl) {
      ui.alert("No folder URL provided. Process aborted.");
      return;
    }
  
    try {
      var folderId = folderUrl.match(/[-\w]{25,}/)[0]; // Extract folder ID from URL
      var folder = DriveApp.getFolderById(folderId);
    } catch (e) {
      ui.alert("Invalid Google Drive folder URL. Please provide a correct URL.");
      return;
    }
  
    // Step 10: Convert each split sheet into CSV and upload
    var splitSheets = ss.getSheets().filter(sheet => sheet.getName().startsWith("Customer_Split_"));
  
    splitSheets.forEach(sheet => {
      var csvData = [];
      var values = sheet.getDataRange().getValues();
  
      values.forEach(row => {
        csvData.push(row.map(value => `"${value}"`).join(",")); // Format as CSV
      });
  
      var csvBlob = Utilities.newBlob(csvData.join("\n"), "text/csv", sheet.getName() + ".csv");
      folder.createFile(csvBlob);
    });
  
    SpreadsheetApp.getUi().alert(`Splitting & Upload Complete! Created ${numSheets} CSV files in the provided Google Drive folder.`);
  }
  
  //Customer tab
  function createCustomerTab() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const salesSheet = ss.getSheetByName("Sales QID Final");
  
    if (!salesSheet) {
      SpreadsheetApp.getUi().alert("Sheet 'Sales QID Final' not found.");
      return;
    }
  
    // Delete existing "Customer" tab if it exists
    let customerSheet = ss.getSheetByName("Customer");
    if (customerSheet) {
      ss.deleteSheet(customerSheet);
    }
  /*customer tab */
    // Create a new "Customer" sheet-----> Step-1
    customerSheet = ss.insertSheet("Customer");
  
    // Get all data from "Sales QID Final"
    const data = salesSheet.getDataRange().getValues();
    const headers = data[0]; // Get headers
  
    // Identify column indexes
    const helperColIndex = headers.indexOf("Helper");
    const nameColIndex = headers.indexOf("Name");
    const memoColIndex = headers.indexOf("Memo");
  
    if (helperColIndex === -1 || nameColIndex === -1 || memoColIndex === -1) {
      SpreadsheetApp.getUi().alert("One or more required columns are missing in 'Sales QID Final' tab.");
      return;
    }
  
    // Filter rows where "Helper" (F) = 1 and extract "Name" & "Memo"
    const filteredData = [["Name", "Memo"]]; // Headers for the "Customer" tab
  
    for (let i = 1; i < data.length; i++) {
      if (data[i][helperColIndex] === 1) {
        filteredData.push([data[i][nameColIndex], data[i][memoColIndex]]);
      }
    }
  
    // If no valid rows found, show alert and stop execution
    if (filteredData.length === 1) {
      SpreadsheetApp.getUi().alert("No data found where 'Helper' is 1.");
      return;
    }
  
    // Write data to "Customer" tab
    customerSheet.getRange(1, 1, filteredData.length, 2).setValues(filteredData);
  
    SpreadsheetApp.getUi().alert("Customer tab has been successfully created.");
  }
  
  // Create a new "Customer" sheet-----> Step-2
  function splitAndUploadCustomerTab() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var customerSheet = ss.getSheetByName("Customer");
  
    if (!customerSheet) {
      SpreadsheetApp.getUi().alert("Sheet 'Customer' not found.");
      return;
    }
  
    var data = customerSheet.getDataRange().getValues();
    if (data.length <= 1) {
      SpreadsheetApp.getUi().alert("No data found in 'Customer' tab.");
      return;
    }
  
    // Extract headers and filter non-null "Name" (A) column values
    var headerRow = data[0];
    var validRows = data.slice(1).filter(row => row[0]); // Filter rows where "Name" (A) is NOT blank
  
    if (validRows.length === 0) {
      SpreadsheetApp.getUi().alert("No non-null 'Name' values found in 'Customer' tab.");
      return;
    }
  
    var maxRowsPerFile = 990; // Split size
    var totalRows = validRows.length;
    var numFiles = Math.ceil(totalRows / maxRowsPerFile);
  
    // 🔹 Ask for Google Drive folder URL
    var ui = SpreadsheetApp.getUi();
    var response = ui.prompt("Enter Google Drive folder URL where CSV files will be uploaded:");
    var folderUrl = response.getResponseText().trim();
  
    if (!folderUrl) {
      ui.alert("No folder URL provided. Process aborted.");
      return;
    }
  
    try {
      var folderId = folderUrl.match(/[-\w]{25,}/)[0]; // Extract folder ID from URL
      var folder = DriveApp.getFolderById(folderId);
    } catch (e) {
      ui.alert("Invalid Google Drive folder URL. Please provide a correct URL.");
      return;
    }
  
    // 🔹 Create and upload CSV files
    for (var i = 0; i < numFiles; i++) {
      let startRow = i * maxRowsPerFile;
      let endRow = Math.min(startRow + maxRowsPerFile, totalRows);
  
      let chunk = validRows.slice(startRow, endRow);
      let csvData = [headerRow].concat(chunk).map(row => row.join(",")).join("\n"); // Format CSV
  
      let fileName = `Customer_Split_${i + 1}.csv`;
      let blob = Utilities.newBlob(csvData, "text/csv", fileName);
      folder.createFile(blob);
    }
  
    ui.alert(`Splitting & Upload Complete! Created ${numFiles} CSV files in the provided Google Drive folder.`);
  }
  function mainFunctionD(){
  
   applyFiltersAndFormatting();
  
  }
  function applyFiltersAndFormatting() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
  
    // Step 1: Apply filter & freeze header row in all tabs
    sheets.forEach(sheet => {
      const lastRow = sheet.getLastRow();
      const lastColumn = sheet.getLastColumn();
  
      if (lastRow > 0 && lastColumn > 0) {
        // Remove existing filters
        if (sheet.getFilter()) {
          sheet.getFilter().remove();
        }
        // Apply new filter to all columns
        sheet.getRange(1, 1, lastRow, lastColumn).createFilter();
        
        // Freeze the first row (header)
        sheet.setFrozenRows(1);
      }
    });
  
    // Step 2: Apply specific formatting to "Sales QID Working" tab
    const salesQIDWorkingSheet = ss.getSheetByName("Sales QID Working");
    if (salesQIDWorkingSheet) {
      const lastRow = salesQIDWorkingSheet.getMaxRows();
  
      // Apply color to "New Description" (Column O)
      salesQIDWorkingSheet.getRange(`O1:O${lastRow}`).setBackground("#FCE7C8");
  
      // Apply color to "COA Working" (Column R)
      salesQIDWorkingSheet.getRange(`R1:R${lastRow}`).setBackground("#FCE7C8");
  
      // Apply color to "Final COA" (Column S)
      salesQIDWorkingSheet.getRange(`S1:S${lastRow}`).setBackground("#FCE7C8");
  
      // Apply color to headers of columns AD to AJ
      salesQIDWorkingSheet.getRange("AD1:AJ1").setBackground("#6A80B9");
    }
  
    // Step 3: Apply specific formatting to "Statement" tab
    const statementSheet = ss.getSheetByName("Statement");
    if (statementSheet) {
      const lastColumn = statementSheet.getLastColumn();
      
      // Ensure K-Z is within the lastColumn range
      const startCol = 11; // Column K (1-based index)
      const endCol = 26; // Column Z (1-based index)
      if (lastColumn >= endCol) {
        statementSheet.getRange(1, startCol, 1, endCol - startCol + 1).setBackground("#143D60");
      }
    }
  
    SpreadsheetApp.flush(); // Ensure all updates are applied immediately
    SpreadsheetApp.getUi().alert("Filters, Freezing, and Formatting applied successfully!");
  }
  
  
  
  
  