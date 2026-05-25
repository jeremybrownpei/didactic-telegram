
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
    activeEndCell.setValue(excelDate);
  }

  //clear timer description
  sheet.getRange("A2").setValue("");
}
