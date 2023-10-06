const priorityArray = ["Critical", "High", "Medium", "Low"];


const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
const sheetName = activeSheet.getName();
const cellSheetNameValue = activeSheet.getRange('A1').getValue();
const cellTaskCount = activeSheet.getRange('A2');

const activeCell = activeSheet.getActiveCell();
const rowIdx = activeCell.getRowIndex();
    const colA = activeSheet.getRange('A:A').getValues();



function handleSheetNameChange() {
  // Logger.log(colA)
  colA.map((tId, idx) => {
      // Logger.log(tId)
      const task = { idvalue: "" };

      if (idx > 1 && tId.length > 0) { //skip first two rows
          let tIdSplit = tId[0].split("");

          for (let i = 0; i < 5; i++) { // account for a 6 digit ID number
              const lastLetter = tIdSplit.pop();
              if (lastLetter == "_" || !parseInt(lastLetter)) { // if last letter is NaN
                  break;
              } else if (!task.idvalue) {
                  task.idvalue = `${lastLetter}`;
              } else {
                  const str = `${lastLetter}${task.idvalue}`
                  task.idvalue = str;
                }
            }
        }
      Logger.log(task.idvalue);

      })
    }
 
  function handleCreateUniqueTaskID() {
    const taskIdCell = activeSheet.getRange(`A${rowIdx}`);

    if (!taskIdCell.getValue()) {
        const taskId = `${cellSheetNameValue}` + "_" + `${cellTaskCount.getValue()}`;

        activeSheet.getRange('A2').setValue(cellTaskCount.getValue() +1); //update task counter

        taskIdCell.setValue(`${taskId}`); //add id to first column
    } 
  }



function onEdit(e) { 
    const activeCellValue = activeCell.getValue();
      if (e && priorityArray.includes(activeCellValue)) {
          handleCreateUniqueTaskID();
        } else {
          handleSheetNameChange()
        }
    }
