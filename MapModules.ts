function mapModulesUG(userWS: string) {
  // Set references to map sheet and module matrix
  const map: Excel.Worksheet = Excel.Workbook.getActiveWorkbook().getWorksheet("Y1+Y2+Y3");
  const matrix: Excel.Worksheet = Excel.Workbook.getActiveWorkbook().getWorksheet(userWS);
  const lastRowMapping: number = map.getRange("C" + map.getUsedRange().getRowCount()).getEnd(ExcelScript.Direction.up).getRow();
  let targetCell: Excel.Range = matrix.getRange("A13");

  // Loop through each cell in column of mapping sheet; adapt range as necessary.
  for (let sourceCell of map.getRange("C12:C" + lastRowMapping).getValues()) {
    if (typeof sourceCell[0] === 'string') { // Check if the value is a string
      // Map the string value to module matrix
      targetCell.setValue(sourceCell[0]);
      
      let outcomeCell: Excel.Range = sourceCell.getOffset(1, 7);
      
      let n: number = 1;
      // Iterates through each 25 learning outcomes in Undergraduate Course
      while (n < 25) {
        const colJLetter: string = outcomeCell.getAddress().split('$')[1] + "3"; // Get column letter of outcomeCell
        const JCell: Excel.Range = map.getRange(colJLetter);

        if (JCell.getText().trim().charAt(0) === "C") {
          switch (outcomeCell.getText()) {
            case "B: Delivered & Assessed by BOTH Examination & Continuous Assessment":
              targetCell.getOffset(0, 2 + n).setValue("B");
              break;
            case "C: Delivered & Assessed by Continuous Assessment":
              targetCell.getOffset(0, 2 + n).setValue("C");
              break;
            case "D: Delivered or Developed but not assessed":
              targetCell.getOffset(0, 2 + n).setValue("D");
              break;
            case "E: Delivered & Assessed by Examination or in-class test":
              targetCell.getOffset(0, 2 + n).setValue("E");
              break;
          }
        } else if (JCell.getText().trim().charAt(0) === "M") {
          n -= 1;
        }
        n += 1;
        
        // Move pointer to next learning outcome (right side)
        outcomeCell = outcomeCell.getOffset(0, 1);
      }
      // Move to the next row in mapping sheet
      targetCell = targetCell.getOffset(1, 0);
    }
  }
  
  // Removing column data from EQF levels below 7 
  for (let outcomeCell of matrix.getRange("D10:U10").getValues()) {
    const columnLetter: string = outcomeCell.getAddress().split('$')[1];
    if (outcomeCell[0].charAt(outcomeCell[0].length - 1) === "*") {
      matrix.getRange(columnLetter + "13:" + columnLetter + "50").clear(ExcelScript.ClearApplyTo.contents);
    }
  }
}
