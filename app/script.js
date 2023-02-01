'use strict';

window.onload = function () {

    //FUNCTION RUNS ON BUTTON CLICK
    document.getElementById("demo").onclick = () => {
       
        //TABLEAU EXTENSION API
        tableau.extensions.initializeAsync().then(function () {
            let dashboard = tableau.extensions.dashboardContent.dashboard;
            
            // CREATE NEW EXCEL "FILE"
            var workbook = XLSX.utils.book_new();

            //PROMISE
            Promise.all([processDashboard(dashboard, workbook)]).then((values) => {
                console.log('in 3rd task');
                // "FORCE DOWNLOAD" XLSX FILE
                var today = new Date();
                var date = today.getFullYear() + '-' + (today.getMonth() + 1) + '-' + today.getDate();
                var time = today.getHours() + "_" + today.getMinutes() + "_" + today.getSeconds();
                var dateTime = date + ' ' + time;

                XLSX.writeFile(workbook, "Excel-Output _" + dateTime + ".xlsx");
            });



        });
    }
}

function fitToColumn(arrayOfArray) {
    // get maximum character of each column
    return arrayOfArray[0].map((a, i) => ({ wch: Math.max(...arrayOfArray.map(a2 => a2[i] ? a2[i].toString().length : 0)) }));
}

//find whether object with specific value is present in array of objects
function getIndex(arr, name) {
    const { length } = arr;
    const id = length + 1;
    return arr.findIndex(el => el.fieldName === name);
}

function processDashboard(dashboard, workbook) {
    //DECLARE REQUIRED OBJECTS FOR STYLEJS
    const DEF_Size14Vert = { font: { sz: 24 }, alignment: { vertical: 'center', horizontal: 'center' } };
    const DEF_FxSz14RgbVert = { font: { name: 'Calibri', sz: 11, color: { rgb: '000000' } }, alignment: { vertical: 'center', horizontal: 'center' } };
    let detailsWorksheet;

    return new Promise(async function (resolve, reject) {
        let arr = dashboard.worksheets;
        let worksheetArr = [];

        let sheetCount = arr.reduce((accumulator, obj) => {
            if (obj.name.includes('Report_Export_Details_D')) {
                return accumulator + 1;
            }
            return accumulator;
        }, 0);

        let checkCount = 0;

        await dashboard.worksheets.forEach(async function (worksheet, key, arr) {
            if (worksheet.name.includes('Report_Export_Details_D')) {
                detailsWorksheet = worksheet;
                await detailsWorksheet.getSummaryDataAsync().then(async function (mydata) {
                    let dashboardData = mydata.data;
                    let dashboardColumns = mydata.columns;

                    // console.log(mydata);
                    let sheetName = dashboardData[0][getIndex(dashboardColumns, 'Sheet name')].value;
                    let reportHeader = dashboardData[0][getIndex(dashboardColumns, 'Report Header')].value;
                    let reportRefreshTime = dashboardData[0][getIndex(dashboardColumns, 'Report Refresh Time')].value;
                    let reportFooter = dashboardData[0][getIndex(dashboardColumns, 'Report Footer')].value;
                    let param1 = getIndex(dashboardColumns, 'Param1') != -1 ? 
                        dashboardData[0][getIndex(dashboardColumns, 'Param1')].value : '';
                    let param2 = getIndex(dashboardColumns, 'Param2') != -1 ? 
                        dashboardData[0][getIndex(dashboardColumns, 'Param2')].value : '';
                    let param3 = getIndex(dashboardColumns, 'Param3') != -1 ? 
                        dashboardData[0][getIndex(dashboardColumns, 'Param3')].value : '';
                    let param4 = getIndex(dashboardColumns, 'Param4') != -1 ? 
                        dashboardData[0][getIndex(dashboardColumns, 'Param4')].value : '';
                    let user = dashboardData[0][getIndex(dashboardColumns, 'User')].value;
                    //let sheetOrder = dashboardData[0][getIndex(dashboardColumns, 'Sheet order')].value;

                    await dashboard.worksheets.forEach(async function (sheet) {
                        if (sheet.name === sheetName) {
                            await sheet.getSummaryDataAsync().then(function (d) {
                                let sheetData = d;
                                checkCount++;
                                // console.log(sheetData);
                                let columnLength = sheetData.columns.length;
                                let columns = sheetData.columns;
                                let result = [];

                                let tt = [];
                                let rr = [];
                                let empt = [];
                                let ii = columnLength % 2;

                                for (let i = 0; i < columnLength; i++) {
                                    if (i == 0) {
                                        tt.push({ v: reportHeader, t: 's', s: { ...DEF_Size14Vert, fill: { fgColor: { rgb: '404040' } }, font: { sz: 22, name: 'Calibri', color: { rgb: 'f1f1f1' } } } });
                                    } else {
                                        tt.push({ v: ' ', t: 's', s: { ...DEF_Size14Vert, fill: { fgColor: { rgb: '404040' } }, font: { sz: 22, name: 'Calibri', color: { rgb: 'f1f1f1' } } } });
                                    }
                                    if (i == columnLength - 2) {
                                        rr.push({ v: `Report executed by ${user} ${reportRefreshTime}`, t: 's', s: { ...DEF_Size14Vert, fill: { fgColor: { rgb: '404040' } }, font: { sz: 11, name: 'Calibri', color: { rgb: 'f1f1f1' } }, alignment: { horizontal: 'right' } } });
                                    } else {
                                        rr.push({ v: ' ', t: 's', s: { ...DEF_Size14Vert, fill: { fgColor: { rgb: '404040' } }, font: { sz: 11, name: 'Calibri', color: { rgb: 'f1f1f1' } }, alignment: { horizontal: 'right' } } });
                                    }
                                    empt.push(" ");
                                }

                                result.push(tt);

                                tt = [];
                                for (let i = 0; i < columnLength; i++) {
                                    if (i == columnLength -2) {
                                        let p = param1 != '' && param2 != '' && param3 != '' && param4 != '' ? `${param1};   ${param2};   ${param3};   ${param4};` : param1 != '' && param2 != '' && param3 != '' ? `${param1};   ${param2};   ${param3};` : param1 != '' && param2 != '' ? `${param1};   ${param2};` : param1 != '' ? `${param1};` : '';
                                        tt.push({ v: p, t: 's', s: { ...DEF_FxSz14RgbVert, border: { right: { style: 'thin', color: { rgb: '000000' } }, left: { style: 'thin', color: { rgb: '000000' } }, bottom: { style: 'thin', color: { rgb: '000000' } }, top: { style: 'thin', color: { rgb: '000000' } } }, font: { sz: 11, name: 'Calibri', bold: true }, alignment: { horizontal: 'right' } } });
                                    } else {
                                        tt.push({ v: '', t: 's', s: { ...DEF_FxSz14RgbVert, border: { right: { style: 'thin', color: { rgb: '000000' } }, left: { style: 'thin', color: { rgb: '000000' } }, bottom: { style: 'thin', color: { rgb: '000000' } }, top: { style: 'thin', color: { rgb: '000000' } } }, font: { sz: 11, name: 'Calibri', bold: true }, alignment: { horizontal: 'right' } } });
                                    }
                                }

                                result.push(empt);
                                result.push(rr);
                                result.push(tt);
                                result.push(empt);
                                result.push(empt);

                                tt = [];
                                for (let i = 0; i < columnLength; i++) {
                                    let colEle = columns[i];
                                    tt.push({ v: colEle.fieldName, t: 's', s: { ...DEF_FxSz14RgbVert, border: { right: { style: 'thin', color: { rgb: '000000' } }, left: { style: 'thin', color: { rgb: '000000' } }, bottom: { style: 'thin', color: { rgb: '000000' } }, top: { style: 'thin', color: { rgb: '000000' } } }, font: { sz: 11, name: 'Calibri', bold: true } } });
                                }

                                result.push(tt);

                                let colData = sheetData.data;
                                for (let i = 0; i < colData.length; i++) {
                                    let arrEle = colData[i];
                                    let tempArr = [];
                                    for (let j = 0; j < arrEle.length; j++) {
                                        tempArr.push({ v: arrEle[j].value == '%null%' ? 'Null' : arrEle[j].value, t: isNaN(arrEle[j].value) ? 's' : 'n', s: { ...DEF_FxSz14RgbVert, border: { right: { style: 'thin', color: { rgb: '000000' } }, left: { style: 'thin', color: { rgb: '000000' } }, bottom: { style: 'thin', color: { rgb: '000000' } }, top: { style: 'thin', color: { rgb: '000000' } } }, alignment: isNaN(arrEle[j].value) ? {horizontal: 'left' } : {horizontal: 'right' }} });
                                    }
                                    result.push(tempArr);
                                }

                                tt = [];
                                for (let i = 0; i < columnLength; i++) {
                                    if (i == 0) {
                                        tt.push({ v: reportFooter, t: 's', s: { ...DEF_Size14Vert, fill: { fgColor: { rgb: '404040' } }, font: { sz: 11, name: 'Calibri', color: { rgb: 'f1f1f1' } }, alignment: { vertical: 'bottom', horizontal: 'center' } } });
                                    } else {
                                        tt.push({ v: ' ', t: 's', s: { ...DEF_Size14Vert, fill: { fgColor: { rgb: '404040' } }, font: { sz: 11, name: 'Calibri', color: { rgb: 'f1f1f1' } }, alignment: { vertical: 'bottom', horizontal: 'center' } } });
                                    }
                                }

                                result.push(empt);
                                result.push(empt);
                                result.push(tt);

                                //CREATE WORKSHEET(S) AND ADD IT TO EXCEL FILE
                                let worksheet = XLSX.utils.aoa_to_sheet(result);

                                let rowFooterMergeStart = 9 + sheetData.totalRowCount;

                                worksheet['!cols'] = fitToColumn(result);
                                worksheet['!rows'] = [{ 'hpt': 40 }];
                                worksheet["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 1, c: columnLength - 1 } },
                                { s: { r: 2, c: columnLength - 2 }, e: { r: 2, c: columnLength - 1 } },
                                { s: { r: 3, c: columnLength - 2 }, e: { r: 3, c: columnLength - 1 } },
                                { s: { r: rowFooterMergeStart, c: 0 }, e: { r: rowFooterMergeStart + 2, c: columnLength - 1 } }
                                ];

                                let obj = {
                                    //index: sheetOrder,
                                    name: sheetName,
                                    worksheet: worksheet
                                }

                                worksheetArr.push(obj);
                                
                                if (sheetCount == checkCount) {
                                    //worksheetArr.sort((a, b) => a.index - b.index);
                                    worksheetArr.forEach((worksheetInfo) => {
                                        //console.log(worksheetInfo.index);
                                        workbook.SheetNames.push("Excel-Output- " + worksheetInfo.name);
                                        workbook.Sheets["Excel-Output- " + worksheetInfo.name] = worksheetInfo.worksheet;
                                    });
                                    console.log('ended');
                                    resolve();
                                }
                            });
                            // return;
                        }
                    });
                });
            }
            // if (Object.is(arr.length -1, key)) {
            //     checkLast = true;
            // }
        });
        // console.log(check);
    });
}