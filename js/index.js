const Excel = require('exceljs');
const { ipcRenderer } = require('electron');
const fs = require('fs');
const moment = require('moment');

let businessReportFile = document.getElementById('businessReportFile');
let kongWeeklyFile = document.getElementById('kongWeeklyFile');
let kongInventoryFile = document.getElementById('kongInventoryFile');
let formulaSportsWeeklyFile = document.getElementById('formulaSportsWeeklyFile');
let formulaSportsInventoryFile = document.getElementById('formulaSportsInventoryFile');
let rapidLossWeeklyFile = document.getElementById('rapidLossWeeklyFile');
let rapidLossInventoryFile = document.getElementById('rapidLossInventoryFile');
let skinPhysicsWeeklyFile = document.getElementById('skinPhysicsWeeklyFile');
let skinPhysicsInventoryFile = document.getElementById('skinPhysicsInventoryFile');
let businessReportSubmit = document.getElementById('businessReportSubmit');
let year, month, saturday, yearAWeekAgo, monthAWeekAgo, sundayAWeekAgo;
//let businessReportTitle, kongWeeklyTitle, formulaSportsWeeklyTitle, rapidLossWeeklyTitle, skinPhysicsWeeklyTitle;

function getFileLocations() {
  return JSON.parse(fs.readFileSync('./assets/files/weekly-locations.json'));
}

function chooseAFile(jsonVariable, filePath, HTMLElementTitleId) {
  let weeklyLocations = getFileLocations();
  weeklyLocations[jsonVariable] = filePath;

  let writeJson = JSON.stringify(weeklyLocations, null, 2);
  fs.writeFileSync('./assets/files/weekly-locations.json', writeJson, (err) => {
    if (err) throw err;
  });

  var splitArray = String(filePath).split('/');
  fileTitle = splitArray[splitArray.length - 1];
  document.getElementById(HTMLElementTitleId).innerHTML = fileTitle;
}

//changes the value of the JSON file to a new String value
function changeFileLocation(jsonVariable, newPath) {
  let changeLocations = JSON.stringify(getFileLocations(), null, 2);
  changeLocations[jsonVariable] = newPath;
  fs.writeFileSync('./assets/files/weekly-locations.json', changeLocations, (err) => {
    if (err) throw err;
  });
}

businessReportSubmit.addEventListener('click', async (event) => {
  let filePaths = getFileLocations();
  let weeklyFilePaths = {};

  for (var file in filePaths) {
    if (!file.includes('Inventory')){
      weeklyFilePaths[file] = filePaths[file];
    }

    if (filePaths[file] === 'No file chosen' || !fs.existsSync(filePaths[file])) {
      alert('Please upload the missing files');
      return;
    }
  }

  for (var file in weeklyFilePaths) {
    if (file !== 'businessReportFilePath' && file === 'kongWeeklyFilePath') {


      const businessReportWorkbook = new Excel.Workbook();
      const businessReportWorksheet = await businessReportWorkbook.csv.readFile(filePaths.businessReportFilePath);

      let SKUColumn = businessReportWorksheet.getColumn('D');
      let rowsToSplice = [];
      let columnsToSplice = [1, 6, 8];
      let companyName = String(file).substr(0, file.length - 14);

      function spliceRowsBusinessReport() {
        businessReportWorksheet.addRow([]);

        for (var i = rowsToSplice.length - 1; i >= 0; i--) {
          businessReportWorksheet.spliceRows(rowsToSplice[i], 1);
        }

        for (var i = columnsToSplice.length - 1; i >= 0; i--) {
          businessReportWorksheet.spliceColumns(columnsToSplice[i], 1);
        }
      }

      function getCurrentDateInfo() {

        let dateForManipulation = new Date();
        
        let dateSaturday = dateForManipulation.getDate() - dateForManipulation.getDay() - 1;
        let dateSundayBefore = dateSaturday - 6;

        let saturdayDate = new Date(dateForManipulation.setDate(dateSaturday));
        saturday = saturdayDate.getDate();
        year = saturdayDate.getFullYear();
        month = saturdayDate.getMonth();

        let dateAWeekAgo = new Date(dateForManipulation.setDate(dateSundayBefore));
        sundayAWeekAgo = dateAWeekAgo.getDate();
        yearAWeekAgo = dateAWeekAgo.getFullYear();
        monthAWeekAgo = dateAWeekAgo.getMonth();
      }

      getCurrentDateInfo();

      async function processWeeklyReport(filePath, inventoryFilePath) {

        const inputWeeklyReportWorkbook = new Excel.Workbook();
        
        await inputWeeklyReportWorkbook.xlsx.readFile(filePath).then(async () => {
          let AUSheetNameToCopy, lastWeekWorksheetName;
          
          const outputWeeklyReportWorkbook = new Excel.Workbook();
          const inputInventoryWorkbook = new Excel.Workbook();
          const inputInventoryWorksheet = await inputInventoryWorkbook.csv.readFile(inventoryFilePath);

          inputWeeklyReportWorkbook.eachSheet(async (worksheet, id) => {
            if (worksheet.name === 'Inventory') {

              const outputWeeklyReportWorksheet = outputWeeklyReportWorkbook.addWorksheet(worksheet.name);

              outputWeeklyReportWorksheet.properties = worksheet.properties;
              outputWeeklyReportWorksheet.state = worksheet.state;

              const columnsToCopy = ['Product',	'Manufacturer',	'SKU',	'ASIN',	'Part Number',	'List Price',	'Amazon Fulfilled',	'Merchant Fulfilled',	'Active',	'Do Not Order',
                'Discontinued',	'Listing Working',	'Fulfillable Qty.',	'Unsellable Qty.',	'Reserved FC Processing Qty',	'Reserved FC Transfer QTY',	'Reserved Customer Order',
                'Ordered Qty',	'Lindon Warehouse Qty.',	'Hebron Warehouse Qty.',	'Toronto Warehouse Qty.',	'Skullcandy Netherlands Warehouse Qty.',	'Venlo Warehouse Qty.',	
                'Top Ideal Hong Kong Warehouse Qty.',	'Thrapston Warehouse Qty.',	'Melbourne Warehouse Qty.',	'UAE BS Warehouse Qty.',	'Dubai Warehouse Qty.',	'Inbound Qty',	
                'Backorder Qty',	'Backorder Count',	'Wholesale Price'];

              outputWeeklyReportWorksheet.getRow(1).model = worksheet.getRow(1).model;

              for (let column in columnsToCopy){
                let columnToCopyIn, columnToCopyFrom; 
                outputWeeklyReportWorksheet.getRow(1).eachCell((cell, colNumber) => {

                  if (cell.value === columnsToCopy[column]){
                    columnToCopyFrom = colNumber; 
                    if (colNumber > 3) {
                      columnToCopyIn = colNumber - 1;
                    } else {
                      columnToCopyIn = colNumber;
                    }
                  }
                });

                inputInventoryWorksheet.getColumn(columnToCopyFrom).eachCell((cell, rowNumber) => {
                  if(rowNumber > 1){
                    outputWeeklyReportWorksheet.getRow(rowNumber).getCell(columnToCopyIn).value = cell.value;
                    console.log(outputWeeklyReportWorksheet.getRow(rowNumber).getCell(columnToCopyIn).value)
                  }
                })
              }

              return;
            } else if(worksheet.name.startsWith('SKU ')) {
              worksheet.removeConditionalFormatting(format => {
                return format.rules[0].type !== 'cellIs';
              });
              worksheet.addConditionalFormatting({
                ref: 'E2:E1000',
                rules: [
                  {
                    type: 'cellIs',
                    operator: 'lessThan',
                    formulae: ['0'],
                    style: {font: {color: {argb: '9c0005'}}},
                  }
                ]
              });
              worksheet.addConditionalFormatting({
                ref: 'E2:E1000',
                rules: [
                  {
                    type: 'cellIs',
                    operator: 'lessThan',
                    formulae: ['0'],
                    style: {fill: {type: 'pattern', pattern: 'solid', bgColor: {argb: 'ffc7ce'}}},
                  }
                ]
              });
              worksheet.addConditionalFormatting({
                ref: 'E2:E1000',
                rules: [
                  {
                    type: 'cellIs',
                    operator: 'greaterThan',
                    formulae: ['0'],
                    style: {font: {color: {argb: '006100'}}},
                  }
                ]
              });
              worksheet.addConditionalFormatting({
                ref: 'E2:E1000',
                rules: [
                  {
                    type: 'cellIs',
                    operator: 'greaterThan',
                    formulae: ['0'],
                    style: {fill: {type: 'pattern', pattern: 'solid', bgColor: {argb: 'c6efce'}}},
                  }
                ]
              });
            } else if (worksheet.name.startsWith('AU ')) {
              AUSheetNameToCopy = id;
              worksheet.removeConditionalFormatting(format => {
                return format.rules[0].type !== 'duplicateValues';
              });
            } else if (worksheet.name === 'Amazon Data') {

              //Create last week's AU Traffic sheet
              lastWeekWorksheetName = 'AU Traffic ' + sundayAWeekAgo + '.' + (monthAWeekAgo + 1) + '.' + String(yearAWeekAgo).slice(0, 2) + ' - ' + saturday + '.' + (month + 1) + '.' + String(year).slice(0, 2);

              const lastWeekReportWorksheet = outputWeeklyReportWorkbook.addWorksheet();
              lastWeekReportWorksheet.model = inputWeeklyReportWorkbook.getWorksheet(AUSheetNameToCopy).model;
              lastWeekReportWorksheet.state = inputWeeklyReportWorkbook.getWorksheet(AUSheetNameToCopy).state;
              lastWeekReportWorksheet.name = lastWeekWorksheetName;

              businessReportWorksheet.spliceRows(2, 0, inputWeeklyReportWorkbook.getWorksheet(AUSheetNameToCopy).getRow(2));
              //Copy the business report data into the AU Traffic sheet
              businessReportWorksheet.eachRow((row, rowNumber) => {
                let newRow = lastWeekReportWorksheet.getRow(rowNumber);
                row.eachCell((cell, colNumber) => {
                  let newCell = newRow.getCell(colNumber);
                  if (colNumber === 6 && rowNumber > 1){
                    newCell.value = parseInt(cell.value.slice(0, -1)) / 100;
                    newCell.numFmt = '#0%'
                  } else if (colNumber === 8 && rowNumber > 2){
                    newCell.value = parseFloat(cell.value.slice(0, -1)) / 100;
                    newCell.numFmt = '#0.00%'
                  } else if (colNumber === 9 && rowNumber > 2){
                    newCell.value = parseFloat(cell.value.slice(1));
                    newCell.numFmt = '$#0.00'
                  } else {
                    newCell.value = cell.value;
                    newCell.numFmt = cell.numFmt;
                  }
                });
              });

              //Copy the Amazon Data sheet
              const outputAmazonDataWorksheet = outputWeeklyReportWorkbook.addWorksheet(worksheet.name);

              outputAmazonDataWorksheet.model = worksheet.model;
              outputAmazonDataWorksheet.state = worksheet.state;

              businessReportWorksheet.spliceColumns(10, 1);
              
              //Copy the business report data into the Amazon Data sheet
              businessReportWorksheet.eachRow((row, rowNumber) => {
                if (rowNumber > 2) {
                  let wholesalePriceValue, costsOfGoodsSoldValue;
                  
                  function calculateWholeSalePrice() {
                    inputWeeklyReportWorkbook.getWorksheet('Inventory').getColumn('C').eachCell((cell, rowNumber) => {
                      if(cell.value === row.getCell('C').value){
                        wholesalePriceValue = inputWeeklyReportWorkbook.getWorksheet('Inventory').getRow(rowNumber).getCell('AF').value;
                      }
                    });
                  }

                  const rowValues = [];
                  rowValues[1] = new Date(year, month, saturday+1)
                  row.eachCell((cell, colNumber) => {
                    if(colNumber === 6) {
                      rowValues[colNumber + 1] = parseInt(cell.value.slice(0, -1)) / 100;
                    } else if (colNumber === 8) {
                      rowValues[colNumber + 1] = parseFloat(cell.value.slice(0, -1)) / 100;
                    } else if (colNumber === 9) {
                      rowValues[colNumber + 1] = parseFloat(cell.value.slice(1));
                    } else {
                      rowValues[colNumber + 1] = cell.value;
                    }
                  });

                  calculateWholeSalePrice();
                  costsOfGoodsSoldValue = wholesalePriceValue * rowValues[8];

                  //rowValues[1] = { formula: "DATEVALUE(\"3/1/2001\")", result: 36951 };
                  rowValues[11] = { formula: "VLOOKUP(D" + rowNumber + ",Inventory!C:AF,30,FALSE)", result: wholesalePriceValue };
                  rowValues[12] = { sharedFormula: 'L2', result: wholesalePriceValue * rowValues[8] };

                  outputAmazonDataWorksheet.addRow(rowValues);
                }
              });

              outputAmazonDataWorksheet.getColumn(1).numFmt = 'd-mmm';

              return;
            }

            //Copy the sheet into the weekly report sheet
            const outputWeeklyReportWorksheet = outputWeeklyReportWorkbook.addWorksheet(worksheet.name);

            outputWeeklyReportWorksheet.model = worksheet.model;
            outputWeeklyReportWorksheet.state = worksheet.state;

            // worksheet.eachRow((row, rowNumber) => {
            //   var newRow = outputWeeklyReportWorksheet.getRow(rowNumber);
            //   row.eachCell((cell, colNumber) => {
            //     var newCell = newRow.getCell(colNumber);
            //     newCell.model = cell.model;
            //     newCell.numFmt = cell.numFmt;
            //   });
            // });
          });

          //Update SKU sheet
          let skuWorksheet = outputWeeklyReportWorkbook.getWorksheet('SKU Week by Week Growth');
          let newColumnNumber = skuWorksheet.getRow(1).cellCount + 1;
          let newColumnLetter = skuWorksheet.getColumn(newColumnNumber).letter;
          let previousColumnLetter = skuWorksheet.getColumn(newColumnNumber-1).letter;

          //Create the first, date cell for the new column
          skuWorksheet.getRow(1).getCell(newColumnNumber).model = skuWorksheet.getRow(1).getCell(newColumnNumber - 1).model;
          skuWorksheet.getRow(1).getCell(newColumnNumber).value = { formula: previousColumnLetter + "1+7",
          result: outputWeeklyReportWorkbook.getWorksheet('Amazon Data').lastRow.getCell(1).value, shareType: 'shared', ref: newColumnLetter + '1:BX1' } 
          skuWorksheet.getRow(1).getCell(newColumnNumber).numFmt = 'd-mmm'

          //Add the new column data
          skuWorksheet.eachRow((row, rowNumber) => {
            if (rowNumber > 1) {
              let newColumnFormulaResult = 0;
              let percentageChangeFormulaResult;

              //Calculate the new column formula result for this row
              outputWeeklyReportWorkbook.getWorksheet(lastWeekWorksheetName).getColumn(3).eachCell((cell, rowNumber) => {
              if (row.getCell(2).value === cell.value) {
                  newColumnFormulaResult = parseInt(outputWeeklyReportWorkbook.getWorksheet(lastWeekWorksheetName).getRow(rowNumber).getCell(7).value);
                }
              });

              //Calculate the percentage change formula result for this row
              if (row.getCell(previousColumnLetter).value.result === 0 || !row.getCell(previousColumnLetter).value.result) {
                percentageChangeFormulaResult = 0;
              } else {
                percentageChangeFormulaResult = (parseInt(newColumnFormulaResult) - parseInt(row.getCell(previousColumnLetter).value.result)) / parseInt(row.getCell(previousColumnLetter).value.result);
              }

              row.getCell(newColumnLetter).value = { formula: "SUMIFS('Amazon Data'!$H:$H,'Amazon Data'!$A:$A,'SKU Week by Week Growth'!$" + newColumnLetter + "$1,'Amazon Data'!$D:$D,'SKU Week by Week Growth'!B" + rowNumber + ")", result: newColumnFormulaResult }
              row.getCell('E').value = { formula: "IFERROR((" + newColumnLetter + rowNumber + "-" + previousColumnLetter + rowNumber + ")/" + previousColumnLetter + rowNumber + ",0)", result: percentageChangeFormulaResult }
            }
          });

          await outputWeeklyReportWorkbook.xlsx.writeFile(companyName + '.xlsx');

        }).then(function () {
          alert(companyName + ' weekly report successfully processed');
        });
      }

      //filter the businessReportWorksheet for each company SKU
      switch (companyName) {
        case 'kong': {
          companyName = 'Kong';
          SKUColumn.eachCell(function (cell, rowNumber) {
            if (!(cell.value.startsWith('AU-KO') || cell.value.startsWith('SKU'))) {
              rowsToSplice.push(rowNumber);
            }
          });
          spliceRowsBusinessReport();
          await processWeeklyReport(filePaths[file], filePaths['kongInventoryFilePath']);
          break;
        }
        case 'formulaSports': {
          companyName = 'Formula Sports';
          SKUColumn.eachCell(function (cell, rowNumber) {
            if (!(cell.value.startsWith('AU-FORM') || cell.value.startsWith('SKU'))) {
              rowsToSplice.push(rowNumber);
            }
          });
          spliceRowsBusinessReport();
          await processWeeklyReport(filePaths[file], filePaths['formulaSportsInventoryFilePath']);
          break;
        }
        case 'rapidLoss': {
          companyName = 'Rapid Loss';
          SKUColumn.eachCell(function (cell, rowNumber) {
            if (!(cell.value.startsWith('AU-RPL') || cell.value.startsWith('SKU'))) {
              rowsToSplice.push(rowNumber);
            }
          });
          spliceRowsBusinessReport();
          await processWeeklyReport(filePaths[file], filePaths['rapidLossInventoryFilePath']);
          break;
        }
        case 'skinPhysics': {
          companyName = 'Skin Physics';
          SKUColumn.eachCell(function (cell, rowNumber) {
            if (!(cell.value.startsWith('AU-SKP') || cell.value.startsWith('SKU'))) {
              rowsToSplice.push(rowNumber);
            }
          });
          spliceRowsBusinessReport();
          await processWeeklyReport(filePaths[file], filePaths['skinPhysicsInventoryFilePath']);
          break;
        }
      }
    }
  }
});

//on selecting the business report file, change file path in the JSON and update element innerHTML
ipcRenderer.on('selected-businessReport', function (event, result) {
  chooseAFile('businessReportFilePath', result.filePaths[0], 'businessReportTitle');
});

//on selecting the kong file, change file path in the JSON and update element innerHTML
ipcRenderer.on('selected-kong', function (event, result) {
  chooseAFile('kongWeeklyFilePath', result.filePaths[0], 'kongWeeklyTitle');
});

ipcRenderer.on('selected-kong-inventory', function (event, result) {
  chooseAFile('kongInventoryFilePath', result.filePaths[0], 'kongInventoryTitle');
});

//on selecting formula sports file, change file path in the JSON and update element innerHTML
ipcRenderer.on('selected-formulaSports', function (event, result) {
  chooseAFile('formulaSportsWeeklyFilePath', result.filePaths[0], 'formulaSportsWeeklyTitle');
});

ipcRenderer.on('selected-formulaSports-inventory', function (event, result) {
  chooseAFile('formulaSportsInventoryFilePath', result.filePaths[0], 'formulaSportsInventoryTitle');
});

//on selecting the rapid loss file, change file path in the JSON and update element innerHTML
ipcRenderer.on('selected-rapidLoss', function (event, result) {
  chooseAFile('rapidLossWeeklyFilePath', result.filePaths[0], 'rapidLossWeeklyTitle');
});

ipcRenderer.on('selected-rapidLoss-inventory', function (event, result) {
  chooseAFile('rapidLossInventoryFilePath', result.filePaths[0], 'rapidLossInventoryTitle');
});

//on selecting the skin physics file, change file path in the JSON and update element innerHTML
ipcRenderer.on('selected-skinPhysics', function (event, result) {
  chooseAFile('skinPhysicsWeeklyFilePath', result.filePaths[0], 'skinPhysicsWeeklyTitle');
});

ipcRenderer.on('selected-skinPhysics-inventory', function (event, result) {
  chooseAFile('skinPhysicsInventoryFilePath', result.filePaths[0], 'skinPhysicsInventoryTitle');
});

function createTitleFromPath(path) {
  var splitArray = path.split('/');
  title = splitArray[splitArray.length - 1];
  return title;
}

//read the weekly report file locations from weekly-locations.json
//set the innerHTML of title elements
//if the files don't exist, update .json to '/' and set the innerHTML to 'No file chosen'
(function initializeWeeklyReportPaths() {
  weeklyLocations = getFileLocations();
  for (var path in weeklyLocations) {
    var fileTitle = String(path).slice(0, path.length - 8) + 'Title';
    if (weeklyLocations[path] !== 'No file chosen') {
      if (path === 'businessReportFilePath' || !fs.existsSync(weeklyLocations[path])) {
        changeFileLocation(path, 'No file chosen');
        document.getElementById(fileTitle).innerHTML = 'No file chosen';
        continue;
      }
      document.getElementById(fileTitle).innerHTML = createTitleFromPath(weeklyLocations[path]);
    } else {
      document.getElementById(fileTitle).innerHTML = weeklyLocations[path];
    }
  }
})();

businessReportFile.addEventListener('click', function (event) {
  ipcRenderer.send('open-business-report');
});

kongWeeklyFile.addEventListener('click', function (event) {
  ipcRenderer.send('open-kong');
});

kongInventoryFile.addEventListener('click', function (event) {
  ipcRenderer.send('open-kong-inventory');
});

formulaSportsWeeklyFile.addEventListener('click', function (event) {
  ipcRenderer.send('open-formulaSports');
});

formulaSportsInventoryFile.addEventListener('click', function (event) {
  ipcRenderer.send('open-formulaSports-inventory');
});

rapidLossWeeklyFile.addEventListener('click', function (event) {
  ipcRenderer.send('open-rapidLoss');
});

rapidLossInventoryFile.addEventListener('click', function (event) {
  ipcRenderer.send('open-rapidLoss-inventory');
});

skinPhysicsWeeklyFile.addEventListener('click', function (event) {
  ipcRenderer.send('open-skinPhysics');
});

skinPhysicsInventoryFile.addEventListener('click', function (event) {
  ipcRenderer.send('open-skinPhysics-inventory');
});




