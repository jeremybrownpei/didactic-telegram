function main(workbook: ExcelScript.Workbook) {
    //Get Sheet
    let sheet = workbook.getWorksheet("Sheet1");

    // Get current date time for Stars and possible ends
    let now = new Date();
    const excelDate = (now.getTime() / (24 * 60 * 60 * 1000)) + 25569.0;

    //Check for an "active" (exntry without and end time) entry C6
    let activeEndCell = sheet.getRange("C6");
    if (activeEndCell.getValueTypes()[0][0] === ExcelScript.RangeValueType.empty) {
        //Set previous 
        //activeEndCell.setValue(now.toLocaleString());
        activeEndCell.setValue(excelDate);
    }


    //Insert a new row    
    //sheet.getRange("A6").select();
    //let selectedRange = workbook.getActiveCell();
    let selectedRange = sheet.getRange('A6');
    selectedRange.getEntireRow().insert(ExcelScript.InsertShiftDirection.down);

    // Get the source cell and target cell
    let sourceCellDescription = workbook.getWorksheet("Sheet1").getRange("A2");
    let targetCellDescription = workbook.getWorksheet("Sheet1").getRange("A6");
    // Copy the value from the source to the target
    targetCellDescription.setValue(sourceCellDescription.getValue());

       
    sheet.getRange("B6").setValue(excelDate);

    // Copy the start new instance button
    let startNewInstanceBtn = sheet.getRange("E1").getValue();
    sheet.getRange("E6").setValue(startNewInstanceBtn);
}
