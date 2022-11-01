'use strict';

window.onload = function () {
    //DECLARE REQUIRED OBJECTS FOR STYLEJS
    const DEF_Size14Vert = { font: { sz: 24 }, alignment: { vertical: 'center', horizontal: 'center' } };
    const DEF_FxSz14RgbVert = { font: { name: 'Calibri', sz: 11, color: { rgb: '000000' } }, alignment: { vertical: 'center', horizontal: 'center' } };

    //FUNCTION RUNS ON BUTTON CLICK
    document.getElementById("demo").onclick = () => {
        let detailsWorksheet;

        //TABLEAU EXTENSION API
        tableau.extensions.initializeAsync().then(function () {
            let dashboard = tableau.extensions.dashboardContent.dashboard;
            dashboard.worksheets.forEach(function (worksheet) {
                if (worksheet.name.includes('Report_Export_Details_D')) {
                    detailsWorksheet = worksheet;
                    detailsWorksheet.getSummaryDataAsync().then(function (mydata) {
                        let dashboardData = mydata.data;
                        let dashboardColumns = mydata.columns;

                        console.log(mydata);
                        let sheetName = dashboardData[0][getIndex(dashboardColumns, 'Sheet name')].value;
                        let reportHeader = dashboardData[0][getIndex(dashboardColumns, 'Report Header')].value;
                        let reportRefreshTime = dashboardData[0][getIndex(dashboardColumns, 'Report Refresh Time')].value;
                        let reportFooter = dashboardData[0][getIndex(dashboardColumns, 'Report Footer')].value;

                        dashboard.worksheets.forEach(function (sheet) {
                            if (sheet.name === sheetName) {
                                sheet.getSummaryDataAsync().then(function (d) {
                                    let sheetData = d;
                                    console.log(sheetData);
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
                                            rr.push({ v: reportRefreshTime, t: 's', s: { ...DEF_Size14Vert, fill: { fgColor: { rgb: '404040' } }, font: { sz: 11, name: 'Calibri', color: { rgb: 'f1f1f1' } }, alignment: { horizontal: 'right' } } });
                                        } else {
                                            rr.push({ v: ' ', t: 's', s: { ...DEF_Size14Vert, fill: { fgColor: { rgb: '404040' } }, font: { sz: 11, name: 'Calibri', color: { rgb: 'f1f1f1' } }, alignment: { horizontal: 'right' } } });
                                        }
                                        empt.push(" ");
                                    }

                                    result.push(tt);
                                    result.push(empt);
                                    result.push(rr);
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
                                            tempArr.push({ v: arrEle[j].value, t: 's', s: { ...DEF_FxSz14RgbVert, border: { right: { style: 'thin', color: { rgb: '000000' } }, left: { style: 'thin', color: { rgb: '000000' } }, bottom: { style: 'thin', color: { rgb: '000000' } }, top: { style: 'thin', color: { rgb: '000000' } } } } });
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

                                    // CREATE NEW EXCEL "FILE"
                                    var workbook = XLSX.utils.book_new(),
                                    worksheet = XLSX.utils.aoa_to_sheet(result);

                                    let rowFooterMergeStart = 8 + sheetData.totalRowCount;

                                    worksheet['!cols'] = fitToColumn(result);
                                    worksheet['!rows'] = [{ 'hpt': 40 }];
                                    worksheet["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 1, c: columnLength - 1 } },
                                    { s: { r: 2, c: columnLength - 2 }, e: { r: 2, c: columnLength - 1 } },
                                    { s: { r: rowFooterMergeStart, c: 0 }, e: { r: rowFooterMergeStart + 2, c: columnLength - 1 } }
                                    ];

                                    workbook.SheetNames.push("Excel-Output");
                                    workbook.Sheets["Excel-Output"] = worksheet;

                                    // "FORCE DOWNLOAD" XLSX FILE
                                    var today = new Date();
                                    var date = today.getFullYear() + '-' + (today.getMonth() + 1) + '-' + today.getDate();
                                    var time = today.getHours() + "_" + today.getMinutes() + "_" + today.getSeconds();
                                    var dateTime = date + ' ' + time;

                                    XLSX.writeFile(workbook, "Excel-Output _" + dateTime + ".xlsx", {password: '1111'});
                                });
                                return;
                            }
                        });
                    });
                }
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