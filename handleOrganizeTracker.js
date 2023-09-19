/*
1. Get all values from sheet.
1a. 
*/ 


function main(workbook: ExcelScript.Workbook) {
    const activeSheet = workbook.getWorksheet("active");
  const letters = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S"];
    let allTableValues = {};
    

// Get all header values
  for (let i = 0; i < letters.length; i++) {
    const letter = letters[i];
    const headerValue = activeSheet.getRange(`${letter}1`).getValue();
    if (headerValue !== "") {
      allTableValues[`${letter}`] = [`${headerValue}`]; 
      // Create key value pair with object notation, value is an array to push row values
    } else {
      break;
    }
  }
  console.log(allTableValues);

// Get all row values, add values to column array 

  for (const key in allTableValues) { //iterate through column headers to access row values
    // console.log(key)
    let consecutiveEmptyRows = 0;
    let rowNumber = 2;
    while (consecutiveEmptyRows < 10) {
      const cell = `${key}` + `${rowNumber}`;
      const rowValue = activeSheet.getRange(`${cell}`).getValue(String);
      if (rowValue !== "") {
        allTableValues[`${key}`].push(`${rowValue}`)
        // console.log(allTableValues[`${key}`])
        consecutiveEmptyRows = 0;
        rowNumber++;
      } else {
        consecutiveEmptyRows++;
      }

    }
  }
  console.log(allTableValues)

}

