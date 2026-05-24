
function main(workbook: ExcelScript.Workbook) {
    //Get Sheet
    let sheet = workbook.getWorksheet("Sheet1");

    // Get current date time for Stars and possible ends
    let now = new Date();

    // Check for un ended timer
    let activeEndCell = sheet.getRange("C6");
    if (activeEndCell.getValueTypes()[0][0] === ExcelScript.RangeValueType.empty) {
        //Set previous 
        activeEndCell.setValue(now.toLocaleString());
    }

    //Set current row description to main description
    // 1. Get the current active cell or selected range
    let activeCell = workbook.getActiveCell();
    // 2. Expand to get the entire row of that selection
    let entireRow = activeCell.getEntireRow();
    // 3. Get a specific cell from that row using a zero-based column index
    // Example: index 1 refers to Column B
    //let targetCell = entireRow.getCell(0, 1); 
    let targetCell = entireRow.getCell(0, 0); 

    sheet.getRange("A2").setValue(targetCell.getValue());

    // Code from Start Timer
    let selectedRange = sheet.getRange('A6');
    selectedRange.getEntireRow().insert(ExcelScript.InsertShiftDirection.down);

    // Get the source cell and target cell
    let sourceCellDescription = workbook.getWorksheet("Sheet1").getRange("A2");
    let targetCellDescription = workbook.getWorksheet("Sheet1").getRange("A6");
    // Copy the value from the source to the target
    targetCellDescription.setValue(sourceCellDescription.getValue());


    sheet.getRange("B6").setValue(now.toLocaleString());

}
