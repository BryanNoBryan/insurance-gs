const name_col = 1
const date_col = 2
const insurance_col = 3 
const deductible_col = 4
const ins_max_remaining_col = 5
const codes_col = 6
const codes_desc_col = 7
const codes_note_col = 8
const cost_col = 9
const coverage_col = 10
const doc_col = 11
const googleDocTemplate_id = "1vSKG2BA0eefxWQwEKTaa82ItTYnfnLCyi8n6pX5LsLo";
const parentFolder_id = "null rn";
const insuranceSheets_id = "144xBPHOj1XupVLsIkhl8RcVOLjxK4VtrzdnnAS2wzvU";

// will be "Careington"
const careington = ['Guardian', 'Cigna', 'AnthemBCBS', 'United Concordia', 'UHC', 'Humana'];
const insurance_types = ['Self Pay', 'Delta Dental Premier', 'Metlife PDP Plus', ...careington, 'Aetna', 'CUNY', 'UFT'];


let deductibleGroups = [
  ['D9310'], //consult
  ['D0160'], //consult
  ['D7210'], //Surgical Extraction
  ['D7140'], //Simple Extraction
  ['D6010'], //Endosteal Implant
];

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Autofill Docs');
  menu.addItem("Create New Docs", "createNewGoogleDocs");
  // menu.addItem("Create Date", "getDate");
  menu.addToUi();
}

function getDate() {
  const now = new Date();
  const date = Utilities.formatDate(now, "America/New_York", "MM/dd/yy");
  return date;
}

function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const row = range.getRow();
  const col = range.getColumn();


  if (col == name_col) {
    rightwardCell = sheet.getRange(row, insurance_col);

    // If the cell below is empty, add dropdown validation
    if (rightwardCell.isBlank()) {
      // Define dropdown options
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(insurance_types, true)
        .setAllowInvalid(false)
        .build();
      rightwardCell.setDataValidation(rule);

      const date = getDate();
      setCellValue(sheet, row, date_col, date);
    }
  }
}

function createNewGoogleDocs() {
  const insurance_spreadsheet = SpreadsheetApp.openById(insuranceSheets_id);
  const sheet = SpreadsheetApp.getActiveSheet();
  const rows = sheet.getDataRange().getValues();
  
  let docs_to_be_generated_list = [];
  

  //get the start of each block, index is row #, which is 1 number behind (js vs. gsheets diff)
  //if has name, and does not have doc, then add to list
  rows.forEach(function(row, index) {
    if (row[name_col - 1] == "") return;
    if (row[doc_col - 1] != "") return;
    docs_to_be_generated_list.push(index + 1);
  });
  //this last element is not actually a doc to be generated, just helps with height calcs

  //generate docs, index is # in list, row is row #, which is 1 number behind (js vs. gsheets diff)
  docs_to_be_generated_list.forEach(function(row, index) {
    let insurance_name = sheet.getRange(row, insurance_col).getValue();
    if (careington.includes(insurance_name)) insurance_name = "Careington";

    let found_end = false;
    let last_row;
    let temp_row = row;
    while (!found_end) {
      temp_row++;
      //if row no name or is not last row, continue to next row
      if (!sheet.getRange(temp_row, name_col).isBlank() || temp_row > sheet.getLastRow()) {
        last_row = temp_row - 1;
        found_end = true;
        break;
      }
    }

    const insurance_sheet = insurance_spreadsheet.getSheetByName(insurance_name);

    const insurance_matrix = insurance_sheet.getDataRange().getValues();

    //info to fill in docs later
    let codes = []
    let codes_descs = []
    let codes_notes = []
    let costs = []
    let coverages = []

    //
    //the rows for this patient, i.e. their codes, etc.
    for (let i = row; i <= last_row; i++) {
      const code = getCellValue(sheet, i, codes_col);
      
      const code_note = getCellValue(sheet, i, codes_note_col);
      codes.push(code);
      
      codes_notes.push(code_note);

      //cost
      let cost;
      for (let j = 0; j < insurance_matrix.length; j++) {
        //hard coded index of misery and suffering
        if (code === insurance_matrix[j][0]) {
          cost = insurance_matrix[j][2];
          const code_desc = insurance_matrix[j][1];
          codes_descs.push(code_desc);
          setCellValue(sheet, i, codes_desc_col, code_desc);
          break;
        }
      }
      if (cost  === undefined) {
        setCellValue(sheet, i, cost_col, "bad dental code");
        console.log(`write row${i} col${cost_col}, undefined`);
      }
      else {
        setCellValue(sheet, i, cost_col, cost);
        costs.push(cost);
      }

      //coverage
      let coverage = getCellValue(sheet, i, coverage_col);
      if (coverage == "") coverage = 0;
      coverages.push(coverage);
    }
    //accumlate date for fillInGoogleDoc function
    //existing
    // codes = []
    // codes_descs = []
    // codes_notes = []
    // costs = []
    // coverages = []

    const name = getCellValue(sheet, row, name_col);
    let date = getCellValue(sheet, row, date_col);
    date = Utilities.formatDate(date, "America/New_York", "MM/dd/yy");
    const insurance = getCellValue(sheet, row, insurance_col);
    const deductible = getCellValue(sheet, row, deductible_col);
    const ins_max_remaining = getCellValue(sheet, row, ins_max_remaining_col);

    console.log(codes);

    fillInGoogleDoc(sheet, row, doc_col, name, date, insurance, deductible, ins_max_remaining, codes, codes_descs, codes_notes, costs, coverages);

    // one full loop of each "block"
  }
  )
}

//
// sheet, row, col is where I'll insert the link into
// codes, codes_notes, cost and coverage are 1D arrays
//
function fillInGoogleDoc(sheet, row, col, name, date, insurance, deductible, ins_max_remaining, codes, codes_descs, codes_notes, costs, coverages) {

  console.log("filling in docs rn");
  
  const googleDocTemplate = DriveApp.getFileById(googleDocTemplate_id);

  const destinationFolder = DriveApp.getFolderById(getSubFolderId(date, parentFolder_id));

  const copy = googleDocTemplate.makeCopy(`${name}`, destinationFolder);
  const doc = DocumentApp.openById(copy.getId());
  const body = doc.getBody();


  // {{first}} {{last}} {{insurance_name}} {{deductible}} {{date}}
  // {{ins_portion}} sum of ins
  // {{pt_portion}} sum of pt 
  // {{ins_max}} ins_max_remaining
  // {{exceeding}} ins_portion - ins_max
  // {{balance}} pt_portion + exceeding

  let ins_money = [];
  let pt_percentages = [];
  let pt_money = [];
  let ins_sum = 0;
  let pt_sum = 0;

  for (let i = 0; i < codes.length; i++) {
    ins_money.push(costs[i] * coverages[i] / 100);
    pt_percentages.push(100 - coverages[i]);
    pt_money.push(costs[i] * pt_percentages[i] / 100);

    console.log(`i${i} ins_money${ins_money[i]} pt_percentage${pt_percentages[i]} pt_money${pt_money[i]} `);
  }

  //
  //deductibleMaxxing START
  //
  let deductibleCounter = deductible;
  // deductibleGroups []
  // remembers what indices(codes) were already counted into the deductible
  let indicesUsed = [];
  // for when the codes are within specified deductibleGroups
  for (let i = 0; i < deductibleGroups.length; i++) {
    if (deductibleCounter == 0) break;
    console.log(`loop 1 i${i}`);
  	for (let j = 0; j < codes.length; j++) {
  		if (indicesUsed.includes(j)) break;
  		console.log(`loop 2 i${i} j${j}`);

  		if (codes[j] == deductibleGroups[i]) {
  			indicesUsed.push(j);
  
  			let cost = costs[j];
  
  			if (cost > deductibleCounter) {
          console.log(`cost > deduct cost ${cost} deductibleCounter ${deductibleCounter}`);
  				let deductedCost = cost - deductibleCounter;
  				ins_money[j] =  deductedCost * (coverages[j] / 100);
  				pt_money[j] = deductedCost * (1 - coverages[j] / 100) + deductibleCounter;
          deductibleCounter = 0;
  			  console.log(`deductedCost ${deductedCost} ins_money ${ins_money[j]} pt_money ${pt_money[j]}`);
        } else {
          console.log(`cost <= deduct cost ${cost} deductibleCounter ${deductibleCounter}`);
  				let deductibleEaten = cost;
  				deductibleCounter -= deductibleEaten;
  				ins_money[j] = 0;
  				pt_money[j] = cost;
          console.log(`deductibleEaten ${deductibleEaten} deductibleCounter ${deductibleCounter} ins_money ${ins_money[j]} pt_money ${pt_money[j]}`);
  			}
        break;
  		}
      console.log(`deductibleCounter ${deductibleCounter} indicesUsed ${indicesUsed}`);
  	}
  }
  // for when there’s still deductible, but no codes in deductibleCodes match
  let maxValue, maximumValueCodeIndex;
  while (maximumValueCodeIndex != -1) {
    console.log('loop 3');
  	if (deductibleCounter == 0) break;

  	maxValue = -1;
  	maximumValueCodeIndex = -1;
  	for (let j = 0; j < codes.length; j++) {
  		if (indicesUsed.includes(j)) break;
      console.log('loop 4');
  		if (costs[j] > maxValue) {
  			maxValue = costs[j];
  			maximumValueCodeIndex = j;
  		}
  	}
  
    let k = maximumValueCodeIndex;
    let cost = costs[k];
  	if (cost > deductibleCounter) {
  		let deductedCost = cost - deductibleCounter;
  		ins_money[k] =  deductedCost * (coverages[k] / 100);
  		pt_money[k] = deductedCost * (1 - coverages[k] / 100) + deductibleCounter;
      deductibleCounter = 0;
  	}
  	else {
  		let deductibleEaten = cost;
  		deductibleCounter -= deductibleEaten;
  		ins_money[k] = 0;
  		pt_money[k] = cost;
  	}
    console.log(`maxValue ${maxValue} maximumValueCodeIndex ${maximumValueCodeIndex} k${k}`);

    indicesUsed.push(maximumValueCodeIndex);
    console.log(`deductibleCounter ${deductibleCounter} indicesUsed ${indicesUsed}`);
  }

  //
  //deductibleMaxxing END
  //

  // ins_sum and pt_sum accumulation
  for  (let i = 0; i < codes.length; i++) {
    ins_sum += ins_money[i];
    pt_sum += pt_money[i];
  }

  //black bar is just
  // ______________________________________________________________________________
  //tables[2] is cells tables[2]
  //repeat table 2 and black bar
  // removal: tables[2].getParent().removeChild(tables[2]);
  
  const tables = body.getTables();
  const tableTemplate = tables[2].copy();
  tables[2].getParent().removeChild(tables[2]);

  //hard coded index of sadness and despair
  let currentIndex = 6;

  for (let i = 0; i < codes.length; i++) {
    const newTable = tableTemplate.copy();
    body.insertTable(currentIndex++, newTable);
    const row = newTable.getRow(0);
    row.getCell(0).clear().setText(`${codes[i]}: ${codes_descs[i]} ${codes_notes[i]}`);
    row.getCell(1).clear().setText('');
    row.getCell(2).clear().setText('');
    row.getCell(3).clear().setText(`$${round2(costs[i])}`);
    row.getCell(4).clear().setText('');
    if (coverages[i] > 0) row.getCell(5).clear().setText(`${coverages[i]}% ($${round2(ins_money[i])})`);
    else row.getCell(5).clear().setText(`N/A*`);
    row.getCell(6).clear().setText('');
    row.getCell(7).clear().setText(`${pt_percentages[i]}% ($${round2(pt_money[i])})`);
    body.insertParagraph(currentIndex++, "______________________________________________________________________________");
  }
  body.insertPageBreak(currentIndex);
  
  let exceedingInsMax;
  let pt_final_cost = pt_sum;
  if (ins_sum > ins_max_remaining) {
    exceedingInsMax = ins_sum - ins_max_remaining;
    pt_final_cost += exceedingInsMax;
  } else {
    exceedingInsMax = 0;
  }

  body.replaceText("{{name}}", name);
  body.replaceText("{{insurance_name}}", insurance);
  if (deductible == "") body.replaceText("{{deductible}}", `N/A`);
  else body.replaceText("{{deductible}}", `${deductible}-${insurance}`);
  body.replaceText("{{date}}", date);
  body.replaceText("{{ins_portion}}", round2(ins_sum));
  body.replaceText("{{pt_portion}}", round2(pt_sum));
  body.replaceText("{{ins_max}}", round2(ins_max_remaining));
  body.replaceText("{{exceeding}}", round2(exceedingInsMax));
  body.replaceText("{{balance}}", round2(pt_final_cost));

  doc.saveAndClose();
  const url = doc.getUrl();
  setCellValue(sheet, row, col, url);
}

function round2(num) {
  num = Number(num);
  if (Number.isInteger(num)) {
    return num.toString();  // no decimals
  } else {
    return num.toFixed(2);  // 2 decimals
  }
}

// 1. Looks inside a parent folder (given its ID).
// 2. Checks if a subfolder with a given date string already exists.
// 3. If it exists → return that folder’s ID.
// 4. If it doesn’t exist → create it, then return the new folder’s ID.
function getSubFolderId(date, parentId) {
  let root = DriveApp.getRootFolder();
  // let parentFolder = DriveApp.getFolderById(parentId);
  let parentFolder = root;
  let subfolders = parentFolder.getFoldersByName(date);
  
  if (subfolders.hasNext()) {
    // Folder already exists
    return subfolders.next().getId();
  } else {
    // Folder does not exist, create it
    let newFolder = parentFolder.createFolder(date);
    return newFolder.getId();
  }
}

function getCellValue(sheet, row, col) {
  const cell = sheet.getRange(row, col);
  return cell.getValue();
}

function setCellValue(sheet, row, col, val) {
  const cell = sheet.getRange(row, col);
  return cell.setValue(val);
}