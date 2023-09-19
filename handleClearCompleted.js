/*
1. Get all values from sheet.
2. Separate the completed & incomplete items.
3. Append completed items to completed sheet.
4. Delete moved items from active sheet.
*/ 


const strTrimmer = function (str) { // Removes extra whitespace before and after text values. Also adjusts double whitespaces to single whitespaces.
  let removeExtraWhitespace = `${str}`.trim().replace(/ {2,}/g, " ");
  return removeExtraWhitespace;
};

function main(workbook: ExcelScript.Workbook) {
  const activeSheet = workbook.getWorksheet("active");
  const letters = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"];
    const allTableValues = {};
    

// Get all header values
  for (let i = 0; i < letters.length; i++) {
    const letter = letters[i];
    const headerValue = activeSheet.getRange(`${letter}1`).getValue();
    if (headerValue !== "") {
      allTableValues[`${letter}`] = [`${headerValue}`]; 
      // Create key value pair with object notation, value is an array to push row values
      } 
      else 
      {
      break;
      }
    }
  // console.log("GET COL VALUES, GET MAX COLS", allTableValues);

// Get all row values, add values to column array 
  for (const key in allTableValues) { //iterate through column headers to access row values
    let consecutiveEmptyRows = 0;
    let rowNumber = 2;
    
    while (consecutiveEmptyRows < 10) {
      const cell = `${key}` + `${rowNumber}`;
      const rowValue = activeSheet.getRange(`${cell}`).getValue();
      if (rowValue !== "") {
        allTableValues[`${key}`].push(`${rowValue}`)
        // console.log(allTableValues[`${key}`])
        consecutiveEmptyRows = 0;
        rowNumber++;
      } 
      else 
      {
        consecutiveEmptyRows++;
       }
      }
    }
  // console.log("GET ROW VALUES", allTableValues)
  // Divide completed/incomplete items into two arrays
  // Get number of rows to create arrays.
  const numberOfRows = getMaxNumberOfRows();
  let completeColIndex = 0;
  function getMaxNumberOfRows(){
    let maxLength_columnIndex = [0, 0]; // [length, index]
    const entries = Object.entries(allTableValues);
    for (let i = 0; i < entries.length; i++) {
      const column = entries[i][1];
      const header = strTrimmer(entries[i][1][0]);
      
      console.log(header);
      if (column.length > maxLength_columnIndex[0]) {
        maxLength_columnIndex = [column.length, i];
      }
    }
    return maxLength_columnIndex;
  };
  // console.log("GET MAX ROWS", numberOfRows)

  
}

