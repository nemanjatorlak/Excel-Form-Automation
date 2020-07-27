const Excel = require('exceljs');
const { ipcRenderer } = require('electron');
const fs = require('fs');
var iconv = require('iconv-lite');
const { isNullOrUndefined } = require('util');

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
let actualPurchasesValueKong = document.getElementById('actualPurchasesValueKong');
let unitsPurchasedValueKong = document.getElementById('unitsPurchasedValueKong');
let actualPurchasesValueFormulaSports = document.getElementById('actualPurchasesValueFormulaSports');
let unitsPurchasedValueFormulaSports = document.getElementById('unitsPurchasedValueFormulaSports');
let actualPurchasesValueRapidLoss = document.getElementById('actualPurchasesValueRapidLoss');
let unitsPurchasedValueRapidLoss = document.getElementById('unitsPurchasedValueRapidLoss');
let actualPurchasesValueSkinPhysics = document.getElementById('actualPurchasesValueSkinPhysics');
let unitsPurchasedValueSkinPhysics = document.getElementById('unitsPurchasedValueSkinPhysics');
let year, month, saturday, yearAWeekAgo, monthAWeekAgo, sundayAWeekAgo;
const monthValues = {
  1: 'Jan',
  2: 'Feb',
  3: 'Mar',
  4: 'Apr',
  5: 'May',
  6: 'June',
  7: 'July',
  8: 'Aug',
  9: 'Sep',
  10: 'Oct',
  11: 'Nov',
  12: 'Dec'
}

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
    if (file !== 'businessReportFilePath') {
      
      function changeEncoding(path) {
        var buffer = fs.readFileSync(path, {encoding: 'binary'});
        var output = iconv.encode(iconv.decode(buffer, 'win1250'), "utf-8");
        fs.writeFileSync(path + 'encoding.csv', output);
      }
      
      changeEncoding(filePaths.businessReportFilePath);

      const businessReportWorkbook = new Excel.Workbook();
      const businessReportWorksheet = await businessReportWorkbook.csv.readFile(filePaths.businessReportFilePath + 'encoding.csv');

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
          let AUSheetNameToCopy, lastWeekWorksheetName, newRowNumberAmazonData = 0;
          
          const outputWeeklyReportWorkbook = new Excel.Workbook();
          const inputInventoryWorkbook = new Excel.Workbook();
          const inputInventoryWorksheet = await inputInventoryWorkbook.csv.readFile(inventoryFilePath);

          inputWeeklyReportWorkbook.eachSheet(async (worksheet, id) => {
            if (worksheet.name === 'Inventory') {

              const outputWeeklyReportWorksheet = outputWeeklyReportWorkbook.addWorksheet(worksheet.name);

              outputWeeklyReportWorksheet.properties = worksheet.properties;
              outputWeeklyReportWorksheet.views = worksheet.views;
              outputWeeklyReportWorksheet.state = worksheet.state;

              // const columnsToCopy = ['Product',	'Manufacturer',	'SKU',	'ASIN',	'Part Number',	'List Price',	'Amazon Fulfilled',	'Merchant Fulfilled',	'Active',	'Do Not Order',
              //   'Discontinued',	'Listing Working',	'Fulfillable Qty.',	'Unsellable Qty.',	'Reserved FC Processing Qty',	'Reserved FC Transfer QTY',	'Reserved Customer Order',
              //   'Ordered Qty',	'Lindon Warehouse Qty.',	'Hebron Warehouse Qty.',	'Toronto Warehouse Qty.',	'Skullcandy Netherlands Warehouse Qty.',	'Venlo Warehouse Qty.',	
              //   'Top Ideal Hong Kong Warehouse Qty.',	'Thrapston Warehouse Qty.',	'Melbourne Warehouse Qty.',	'UAE BS Warehouse Qty.',	'Dubai Warehouse Qty.',	'Inbound Qty',	
              //   'Backorder Qty',	'Backorder Count',	'Wholesale Price'];

              outputWeeklyReportWorksheet.getRow(1).model = worksheet.getRow(1).model;

              //Copy the data from the inventory file into the inventory sheet
              //The FNSKU column and all columns after Wholesale Price are not copied
              outputWeeklyReportWorksheet.getRow(1).eachCell((cell, colNumber) => {
                let columnToCopyIn, columnToCopyFrom; 
                let outputCell = cell, outputColNumber = colNumber;
                inputInventoryWorksheet.getRow(1).eachCell((cell, colNumber) => {
                  if (outputCell.value === cell.value) {
                    columnToCopyIn = outputColNumber;
                    columnToCopyFrom = colNumber;

                    return;
                  }
                });

                inputInventoryWorksheet.getColumn(columnToCopyFrom).eachCell((cell, rowNumber) => {
                  if (cell.value === null) {
                    return;
                  } else if (rowNumber > 1) {
                    if (inputInventoryWorksheet.getRow(1).getCell(columnToCopyFrom).value === 'List Price' || inputInventoryWorksheet.getRow(1).getCell(columnToCopyFrom).value === 'Wholesale Price'){
                      outputWeeklyReportWorksheet.getRow(rowNumber).getCell(columnToCopyIn).value = parseFloat(cell.value.slice(1));
                      outputWeeklyReportWorksheet.getRow(rowNumber).getCell(columnToCopyIn).numFmt = '$#0.00';
                    } else {
                    outputWeeklyReportWorksheet.getRow(rowNumber).getCell(columnToCopyIn).value = cell.value;
                    }
                  }
                });

                if (outputWeeklyReportWorksheet.getColumn(columnToCopyIn).width !== undefined) {
                  outputWeeklyReportWorksheet.getColumn(columnToCopyIn).width = worksheet.getColumn(columnToCopyIn).width;
                } 
              });

            return;
            } else if (worksheet.name.startsWith('SKU ')) {
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
              lastWeekReportWorksheet.state = inputWeeklyReportWorkbook.getWorksheet(AUSheetNameToCopy).state;
              lastWeekReportWorksheet.name = lastWeekWorksheetName;
              //Copy the first two rows from the last week AU Traffic sheet
              businessReportWorksheet.spliceRows(2, 0, [, 'TOTALS']);
              let AUCellF2Summ = 0, AUCellF2Divider = 0, AUCellG2Sum = 0, AUCellFResult = 0, businessReportEmpty = true;
             
              if (businessReportWorksheet.getRow(3).getCell(1).value !== null ) {
                  businessReportEmpty = false;
              }
              
              //Copy the business report data into the AU Traffic sheet
              businessReportWorksheet.eachRow((row, rowNumber) => {
                if (rowNumber <= 2) {
                  lastWeekReportWorksheet.getRow(rowNumber).model = inputWeeklyReportWorkbook.getWorksheet(AUSheetNameToCopy).getRow(rowNumber).model;
                } else {
                  let newRow = lastWeekReportWorksheet.getRow(rowNumber);
                  row.eachCell((cell, colNumber) => {
                    let newCell = newRow.getCell(colNumber);
                    if (colNumber === 6 && rowNumber > 2){
                      let newCellValue = parseInt(cell.value.slice(0, -1)) / 100;
                      newCell.value = newCellValue;
                      AUCellF2Summ += newCellValue;
                      AUCellF2Divider ++;
                      newCell.numFmt = '#0%'
                    } else if (colNumber === 7 && rowNumber > 2){
                      AUCellG2Sum += cell.value;
                      newCell.value = cell.value;
                      newCell.numFmt = cell.numFmt;
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
                }
              });

              if (businessReportEmpty === false){
                AUCellFResult = AUCellF2Summ/AUCellF2Divider;
              }

              lastWeekReportWorksheet.getRow(2).getCell(6).value = { formula: "AVERAGE(F3:F1048576)", result: AUCellFResult };
              lastWeekReportWorksheet.getRow(2).getCell(7).value = { formula: "SUM(G3:G1048576)", result: AUCellG2Sum };

              lastWeekReportWorksheet.addConditionalFormatting({
                ref: 'F2:F1200',
                rules: [
                  {
                    type: 'colorScale',
                    cfvo: [{type: "min", value: 0}, {type: "percentile", value: 50}, {type: "max", value: 0}],
                    color: [{argb: "FFF8696B"}, {argb: "FFFFEB84"}, {argb: "FF63BE7B"}],
                  }
                ]
              });

              //Set column widths
              inputWeeklyReportWorkbook.getWorksheet(AUSheetNameToCopy).getRow(1).eachCell((cell, colNumber) => {
                lastWeekReportWorksheet.getColumn(colNumber).width = inputWeeklyReportWorkbook.getWorksheet(AUSheetNameToCopy).getColumn(colNumber).width;
              })

              //Copy the Amazon Data sheet
              const outputAmazonDataWorksheet = outputWeeklyReportWorkbook.addWorksheet(worksheet.name);

              outputAmazonDataWorksheet.model = worksheet.model;
              outputAmazonDataWorksheet.state = worksheet.state;
              
              outputAmazonDataWorksheet.getColumn('A').eachCell((cell, rowNumber) => {
                newRowNumberAmazonData++;
              });

              businessReportWorksheet.spliceColumns(10, 1);
              
              //Copy the business report data into the Amazon Data sheet
              //If there are no fields to copy, add a row with a date in the first column
              if (businessReportEmpty === true) {
                const rowValues = [];
                rowValues[1] = new Date(year, month, saturday, 2, 0, 0);
                rowValues[7] = 0; 
                rowValues[10] = 0;
                rowValues[11] = { formula: "IFERROR(VLOOKUP(D" + newRowNumberAmazonData + ",Inventory!C:AF,30,FALSE),0)", result: 0 };
                rowValues[12] = { formula: 'K' + newRowNumberAmazonData + '*H' + newRowNumberAmazonData, result: 0 };

                outputAmazonDataWorksheet.spliceRows(newRowNumberAmazonData, 0, rowValues);
              } else {
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
                    rowValues[1] = new Date(year, month, saturday, 2, 0, 0);
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
                    rowValues[11] = { formula: "IFERROR(VLOOKUP(D" + newRowNumberAmazonData + ",Inventory!C:AF,30,FALSE),0)", result: wholesalePriceValue };
                    rowValues[12] = { formula: 'K' + newRowNumberAmazonData + '*H' + newRowNumberAmazonData, result: costsOfGoodsSoldValue };

                    outputAmazonDataWorksheet.spliceRows(newRowNumberAmazonData, 1, rowValues);
                    newRowNumberAmazonData++;
                  }
                });
              }

              outputAmazonDataWorksheet.getColumn(1).numFmt = 'd-mmm';
              outputAmazonDataWorksheet.getColumn(7).numFmt = '#0%';
              outputAmazonDataWorksheet.getColumn(9).numFmt = '#0.00%';
              //Accounting number format
              outputAmazonDataWorksheet.getColumn(10).numFmt = '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)';
              outputAmazonDataWorksheet.getColumn(11).numFmt = '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)';
              outputAmazonDataWorksheet.getColumn(12).numFmt = '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)';
              // outputAmazonDataWorksheet.getColumn(10).numFmt = '$#0.00';
              // outputAmazonDataWorksheet.getColumn(11).numFmt = '$#0.00';
              // outputAmazonDataWorksheet.getColumn(12).numFmt = '$#0.00';

              return;
            }

            //Copy the sheet into the weekly report sheet
            const outputWeeklyReportWorksheet = outputWeeklyReportWorkbook.addWorksheet(worksheet.name);

            outputWeeklyReportWorksheet.model = worksheet.model;
            outputWeeklyReportWorksheet.state = worksheet.state;
          });

          //Update the SKU sheet
          let skuWorksheet = outputWeeklyReportWorkbook.getWorksheet('SKU Week by Week Growth');
          let newColumnNumber = skuWorksheet.getRow(1).cellCount + 1;
          let newColumnLetter = skuWorksheet.getColumn(newColumnNumber).letter;
          let previousColumnLetter = skuWorksheet.getColumn(newColumnNumber-1).letter;

          //Create the first, date cell for the new column
          skuWorksheet.getRow(1).getCell(newColumnNumber).model = skuWorksheet.getRow(1).getCell(newColumnNumber - 1).model;
          skuWorksheet.getRow(1).getCell(newColumnNumber).value = { formula: previousColumnLetter + "1+7",
          result: outputWeeklyReportWorkbook.getWorksheet('Amazon Data').getRow(newRowNumberAmazonData).getCell(1).value } 
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


          //Update the Overview sheet
          const overviewOutputWorksheet = outputWeeklyReportWorkbook.getWorksheet('Overview');
          let newRowNumberOverview = 1;
          overviewOutputWorksheet.getColumn('A').eachCell((cell, rowNumber) => {
            newRowNumberOverview++;
          });
          let newRowValues = [];
          let overviewNewRowDateValue = monthValues[monthAWeekAgo + 1] + ' ' + sundayAWeekAgo + ' - ' + monthValues[month + 1] + ' ' + saturday;
          let columnSevenResult = 0, columnEightResult = 0, columnNineResult = 0, columnElevenResult = 0, amazonDataFirstRow = 1, amazonDataLastRow = newRowNumberAmazonData,
          cellO3 = 0, cellO4 = 0, cellO5 = 0, cellO6 = 0, cellO7 = 0, cellO7Summ = 0, cellO7Divider = 1, cellO8 = 0;

          newRowValues[1] = overviewNewRowDateValue;
          newRowValues[2] = '';
          newRowValues[3] = '';
          newRowValues[4] = '';
          newRowValues[5] = '';
          newRowValues[10] = '';

          //Include the value for Actual Purchases and Units Purchased in the Overview sheet if they have been entered in the input fields
          if (companyName === 'Kong') {
              if (actualPurchasesValueKong.value) {
                newRowValues[2] = parseFloat(actualPurchasesValueKong.value);
              }
              if (unitsPurchasedValueKong.value) {
                newRowValues[4] = parseFloat(unitsPurchasedValueKong.value);
              }
            } else if (companyName === 'Formula Sports') {
              if (actualPurchasesValueFormulaSports.value) {
                newRowValues[2] = parseFloat(actualPurchasesValueFormulaSports.value);
              }
              if (unitsPurchasedValueFormulaSports.value) {
                newRowValues[4] = parseFloat(unitsPurchasedValueFormulaSports.value);
              }
            } else if (companyName === 'Rapid Loss') {
              if (actualPurchasesValueRapidLoss.value) {
                newRowValues[2] = parseFloat(actualPurchasesValueRapidLoss.value);
              }
              if (unitsPurchasedValueRapidLoss.value) {
                newRowValues[4] = parseFloat(unitsPurchasedValueRapidLoss.value);
              }
            } else if (companyName === 'Skin Physics') {
              if (actualPurchasesValueSkinPhysics.value) {
                newRowValues[2] = parseFloat(actualPurchasesValueSkinPhysics.value);
              }
              if (unitsPurchasedValueSkinPhysics.value) {
                newRowValues[4] = parseFloat(unitsPurchasedValueSkinPhysics.value);
              }
            }

          newRowValues[6] = overviewNewRowDateValue;

          //Calculate the Sessions value for the formula field
          outputWeeklyReportWorkbook.getWorksheet(lastWeekWorksheetName).getColumn('D').eachCell((cell, rowNumber) => {
            if (rowNumber > 2){
              columnSevenResult += cell.value;
            }
          });

          newRowValues[7] = { formula: "SUM('" + lastWeekWorksheetName + "'!D:D)", result: columnSevenResult };

          //Calculate the AU Traffic Actual Sales value for the formula field
          outputWeeklyReportWorkbook.getWorksheet(lastWeekWorksheetName).getColumn('I').eachCell((cell, rowNumber) => {
            if (rowNumber > 2){
              columnEightResult += cell.value;
            }
          });

          newRowValues[8] = { formula: "SUM('" + lastWeekWorksheetName + "'!I:I)", result: columnEightResult };

          //Calculate the Amazon Data Actual Sales value for the formula field
          outputWeeklyReportWorkbook.getWorksheet('Amazon Data').getColumn('L').eachCell((cell, rowNumber) => {
            if (cell.value === null) {
              return;
            }
            const cellValue = outputWeeklyReportWorkbook.getWorksheet('Amazon Data').getRow(rowNumber).getCell('A').value;
            const timeValue = new Date(cellValue).getTime();
            if (timeValue === new Date(year, month, saturday, 2, 0, 0).getTime()){
              if (amazonDataFirstRow === 1) {
                amazonDataFirstRow = rowNumber;
              }
              if (cell.value.result) {
                columnNineResult += cell.value.result;
              }
            }
          });

          newRowValues[9] = { formula: "SUM('Amazon Data'!L" + amazonDataFirstRow + ":L" + amazonDataLastRow + ")", result: columnNineResult };

          //Calculate the Amazon Units Sold value for the formula field
          columnElevenResult = outputWeeklyReportWorkbook.getWorksheet(lastWeekWorksheetName).getRow(2).getCell('G').value.result;

          newRowValues[11] = { formula: "'" + lastWeekWorksheetName + "'!G2", result: columnElevenResult};

          overviewOutputWorksheet.spliceRows(newRowNumberOverview, 1, newRowValues)
          //overviewOutputWorksheet.addRow(newRowValues);

          overviewOutputWorksheet.getRow(newRowNumberOverview).eachCell((cell, colNumber) => {
            cell.style = overviewOutputWorksheet.getRow(newRowNumberOverview - 1).getCell(colNumber).style;
            cell.numFmt = overviewOutputWorksheet.getRow(newRowNumberOverview - 1).getCell(colNumber).numFmt;
          })

          //Calculate column O cell values
          cellO3 = new Date(year, month, saturday, 2, 0, 0);
          
          //Calculate the cell O4 value for the formula field
          outputWeeklyReportWorkbook.getWorksheet('Amazon Data').getColumn('E').eachCell((cell, rowNumber) => {
            const cellValue = outputWeeklyReportWorkbook.getWorksheet('Amazon Data').getRow(rowNumber).getCell('A').value;
            const timeValue = new Date(cellValue).getTime();
            if (timeValue === new Date(year, month, saturday, 2, 0, 0).getTime()){
              cellO4 += cell.value;
            }
          });

          //Calculate the cell O5 value for the formula field
          outputWeeklyReportWorkbook.getWorksheet('Amazon Data').getColumn('H').eachCell((cell, rowNumber) => {
            const cellValue = outputWeeklyReportWorkbook.getWorksheet('Amazon Data').getRow(rowNumber).getCell('A').value;
            const timeValue = new Date(cellValue).getTime();
            if (timeValue === new Date(year, month, saturday, 2, 0, 0).getTime()){
              cellO5 += cell.value;
            }
          });

          //Calculate the cell O6 value for the formula field
          outputWeeklyReportWorkbook.getWorksheet('Amazon Data').getColumn('J').eachCell((cell, rowNumber) => {
            const cellValue = outputWeeklyReportWorkbook.getWorksheet('Amazon Data').getRow(rowNumber).getCell('A').value;
            const timeValue = new Date(cellValue).getTime();
            if (timeValue === new Date(year, month, saturday, 2, 0, 0).getTime()){
              cellO6 += cell.value;
            }
          });

          //Calculate the cell O7 value for the formula field
          outputWeeklyReportWorkbook.getWorksheet('Amazon Data').getColumn('G').eachCell((cell, rowNumber) => {
            const cellValue = outputWeeklyReportWorkbook.getWorksheet('Amazon Data').getRow(rowNumber).getCell('A').value;
            const timeValue = new Date(cellValue).getTime();
            if (timeValue === new Date(year, month, saturday, 2, 0, 0).getTime()){
              cellO7Summ += cell.value;
              cellO7Divider ++;
            }
          });

          cellO7 = cellO7Summ/cellO7Divider;

          if (columnNineResult !== 0) {
            cellO8 = (columnEightResult - columnNineResult)/columnEightResult
          }

          overviewOutputWorksheet.getCell('O3').value = cellO3;
          overviewOutputWorksheet.getCell('O4').value = { formula: "SUMIFS('Amazon Data'!E:E,'Amazon Data'!A:A,Overview!O3)", result: cellO4 }
          overviewOutputWorksheet.getCell('O5').value = { formula: "SUMIFS('Amazon Data'!H:H,'Amazon Data'!A:A,Overview!O3)", result: cellO5 }
          overviewOutputWorksheet.getCell('O6').value = { formula: "SUMIFS('Amazon Data'!J:J,'Amazon Data'!A:A,Overview!O3)", result: cellO6 }
          overviewOutputWorksheet.getCell('O7').value = { formula: "AVERAGEIFS('Amazon Data'!G:G,'Amazon Data'!A:A,Overview!O3)", result: cellO7 }
          overviewOutputWorksheet.getCell('O8').value = { formula: "(H" + newRowNumberOverview + "-I" + newRowNumberOverview + ")/H" + newRowNumberOverview, result: cellO8 }

          //Delete the helper businessReport file
          fs.unlinkSync(filePaths.businessReportFilePath + 'encoding.csv');

          let title = companyName + ' AU Weekly Report ' + sundayAWeekAgo + '.' + (monthAWeekAgo + 1) + '.' + String(yearAWeekAgo).slice(0, 2) + ' - ' + saturday + '.' + (month + 1) + '.' + String(year).slice(0, 2);
          await outputWeeklyReportWorkbook.xlsx.writeFile(title + '.xlsx', {});

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

