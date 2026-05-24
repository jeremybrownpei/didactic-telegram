
function main(workbook: ExcelScript.Workbook) {    
  //Get Sheet
  let sheet = workbook.getWorksheet("Sheet1");

  // Get current date time for Stars and possible ends
  let now = new Date();

  //Check for an "active" (exntry without and end time) entry C6
  let activeEndCell = sheet.getRange("C6");
  if (activeEndCell.getValueTypes()[0][0] === ExcelScript.RangeValueType.empty) {
    //Set previous 
    activeEndCell.setValue(now.toLocaleString());
  }

  //clear timer description
  sheet.getRange("A2").setValue("");
}
