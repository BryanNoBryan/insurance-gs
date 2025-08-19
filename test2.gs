const first_name_col = 1
const last_name_col = 2
const date_col = 3
const insurance_col = 4
const deductible_col = 5
const ins_max_remaining_col = 6
const codes_col = 7
const codes_desc_col = 8
const codes_note_col = 9
const cost_col = 10
const coverage_col = 11
const coinsurance_col = 12
const doc_col = 13
const googleDocTemplate_id = "1vSKG2BA0eefxWQwEKTaa82ItTYnfnLCyi8n6pX5LsLo";
const parentFolder_id = "11wD1TOiGDbAhRKoYzmwayZqbuDqreQQi";
const insuranceSheets_id = "144xBPHOj1XupVLsIkhl8RcVOLjxK4VtrzdnnAS2wzvU";

// will be "Careington"
const careington = ['Guardian', 'Cigna', 'AnthemBCBS', 'United Concordia', 'UHC', 'Humana'];
const insurance_types = ['Self Pay', 'Delta Dental Premier', 'Metlife PDP Plus', ...careington, 'Aetna', 'CUNY', 'UFT'];

let applyDeductible = false;
let consult_group = ['', ''];
let ext_implant_group = ['', ''];
let consult_cost = 0;
let consult_code = '';

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


  if (col == first_name_col || col == last_name_col) {
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
    if (row[first_name_col - 1] == "" && row[last_name_col - 1] == "") return;
    if (row[doc_col - 1] != "") return;
    docs_to_be_generated_list.push(index + 1);
  });
  //this last element is not actually a doc to be generated, just helps with height calcs

  console.log(docs_to_be_generated_list);

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
      if (!sheet.getRange(temp_row, first_name_col).isBlank() && !sheet.getRange(temp_row, last_name_col).isBlank() || temp_row > sheet.getLastRow()) {
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
      console.log(`i(row):${i} code ${code}`);
      codes.push(code);
      
      codes_notes.push(code_note);
      console.log(`i${i} code${code} codes${codes}`);

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

      //coverage + coinsurance
      let coverage = getCellValue(sheet, i, coverage_col);
      if (coverage == "") coverage = 0;
      coverages.push(coverage);
      const coinsurance = round2(cost * (1 - coverage / 100));
      setCellValue(sheet, i, coinsurance_col, coinsurance);
      console.log(`coverage ${coverage} coinsurance ${coinsurance}`);
      console.log(`coinsurance write row${i} col${coinsurance_col}, cost${coinsurance}`);
    }
    //accumlate date for fillInGoogleDoc function
    //existing
    // codes = []
    // codes_descs = []
    // codes_notes = []
    // costs = []
    // coverages = []
    const first_name = getCellValue(sheet, row, first_name_col);
    const last_name = getCellValue(sheet, row, last_name_col);
    let date = getCellValue(sheet, row, date_col);
    date = Utilities.formatDate(date, "America/New_York", "MM/dd/yy");
    const insurance = getCellValue(sheet, row, insurance_col);
    const deductible = getCellValue(sheet, row, deductible_col);
    const ins_max_remaining = getCellValue(sheet, row, ins_max_remaining_col);

    console.log(codes);
    console.log('1ins_max_remaining ' + ins_max_remaining);

    fillInGoogleDoc(sheet, row, doc_col, first_name, last_name, date, insurance, deductible, ins_max_remaining, codes, codes_descs, codes_notes, costs, coverages);

    // one full loop of each "block"
  }
  )
}

//
// sheet, row, col is where I'll insert the link into
// codes, codes_notes, cost and coverage are 1D arrays
//
function fillInGoogleDoc(sheet, row, col, first_name, last_name, date, insurance, deductible, ins_max_remaining, codes, codes_descs, codes_notes, costs, coverages) {

  console.log("filling in docs rn");
  console.log(codes);
  
  const googleDocTemplate = DriveApp.getFileById(googleDocTemplate_id);

  const destinationFolder = DriveApp.getFolderById(getSubFolderId(date, parentFolder_id));

  const copy = googleDocTemplate.makeCopy(`${first_name} ${last_name}`, destinationFolder);
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
    if (codes[i] == 'D9310' || codes[i] == 'D0160') {
      consult_code = code[i];
      consult_cost = costs[i];
      costs[i] -= deductible;
    } 

    ins_money.push(costs[i] * coverages[i] / 100);
    pt_percentages.push(100 - coverages[i]);
    if (codes[i] == 'D9310'|| codes[i] == 'D0160') {pt_money.push(costs[i] * pt_percentages[i] / 100 + deductible);}
    else {pt_money.push(costs[i] * pt_percentages[i] / 100);}
    ins_sum += ins_money[i];
    pt_sum += pt_money[i];
    console.log(`i${i} ins_money${ins_money[i]} pt_percentage${pt_percentages[i]} pt_money${pt_money[i]} ins_sum${ins_sum} pt_sum${pt_sum}`);
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
    if (codes[i] == 'D9310' || codes[i] == 'D0160') {
      costs[i] = consult_cost;
      console.log('consult_cost ' + consult_cost);
    }
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

  console.log('2ins_max_remaining ' + ins_max_remaining);
  
  let exceeding;
  let pt_final_cost = pt_sum;
  if (ins_sum > ins_max_remaining) {
    let exceeding = ins_sum - ins_max_remaining;
    pt_final_cost += exceeding;
    ins_sum = ins_max_remaining;
  } else {
    exceeding = 0;
  }


  // let exceeding = ins_sum - ins_max_remaining < 0 ? 0 : ins_sum - ins_max_remaining;
  // // let pt_final_cost = pt_sum + exceeding;
  // // console.log('exceeding ' + exceeding);
  // // console.log('pt_final_cost ' + pt_final_cost);
  // // all out of pocket basically
  // let pt_final_cost = pt_sum;
  // if (ins_sum < deductible) {
  //   pt_final_cost += ins_sum;
  //   ins_sum = 0;
  // } else {
  //   ins_sum -= deductible;
  //   pt_final_cost += deductible;
  // }

  body.replaceText("{{first}}", first_name);
  body.replaceText("{{last}}", last_name);
  body.replaceText("{{insurance_name}}", insurance);
  if (deductible == "") body.replaceText("{{deductible}}", `N/A`);
  else body.replaceText("{{deductible}}", `${deductible}-${insurance}`);
  body.replaceText("{{date}}", date);
  body.replaceText("{{ins_portion}}", round2(ins_sum));
  body.replaceText("{{pt_portion}}", round2(pt_sum));
  body.replaceText("{{ins_max}}", round2(ins_max_remaining));
  body.replaceText("{{exceeding}}", round2(exceeding));
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