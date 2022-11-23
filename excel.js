let myData = [];
let myHeader = [];
let myFooter = [];

function getData() {
    $.ajax({
        url: 'https://randomuser.me/api/?results=10',
        dataType: 'json',
        success: function (data) {
            console.log('getData', data.results);
            showData(data.results);
        },
    });
}

function showData(data) {
    // console.log(data);
    myData = data.map((d) => {
        const sign = Math.random() < 0.5 ? -1 : 1;
        return {
            firstName: d?.name?.first,
            lastName: d?.name?.last,
            email: d?.email,
            phone: d?.phone,
            income: +(Math.random() * 1000).toFixed(2) * sign,
        };
    });
    console.log('myData', myData);
    let html = '<tr><td>Tên</td><td>Họ</td><td>Email</td><td>Phone</td><td>Income</td></tr>';
    let total = 0;
    $.each(myData, function (key, value) {
        html += '<tr>';
        html += '<td>' + value?.firstName + '</td>';
        html += '<td>' + value?.lastName + '</td>';
        html += '<td>' + value?.email + '</td>';
        html += '<td>' + value?.phone + '</td>';
        html += '<td align="right">' + value?.income + '</td>';
        html += '</tr>';
        total += +value?.income;
    });
    html += '<tr>';
    html += '<td colspan="4">Total</td>';
    html += '<td align="right">' + +total.toFixed(2) + '</td>';
    html += '</tr>';
    myFooter = ['Total', '', '', '', +total.toFixed(2)];
    // console.log('html', html);
    $('table tbody').html(html);
}

async function exportToExcel(fileName, sheetName, report) {
    if (!myData || myData.length === 0) {
        console.error('Chưa có dữ liệu');
        return;
    }
    console.log('exportToExcel', myData);

    if (report !== '') {
        myHeader = ['Tên', 'Họ', 'Email', 'Phone', 'Income'];
        exportToExcelPro('Users', 'Users', report, myHeader, myFooter, [
            { width: 15 },
            { width: 15 },
            { width: 30 },
            { width: 20 },
            { width: 20 },
        ]);
        return;
    }

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet(sheetName);
    const header = Object.keys(myData[0]);
    console.log('header', header);
    ws.addRow(header);
    myData.forEach((rowData) => {
        console.log('rowData', rowData);
        row = ws.addRow(Object.values(rowData));
    });

    const buf = await wb.xlsx.writeBuffer();
    saveAs(new Blob([buf]), `${fileName}.xlsx`);
}

async function exportToExcelPro(fileName, sheetName, report, myHeader, myFooter, widths) {
    if (!myData || myData.length === 0) {
        console.error('ChÆ°a cĂ³ data');
        return;
    }
    console.log('exportToExcel', myData);

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet(sheetName);
    const columns = myHeader?.length;
    const title = {
        border: true,
        money: false,
        height: 100,
        font: { size: 30, bold: false, color: { argb: 'FFFFFF' } },
        alignment: { horizontal: 'center', vertical: 'middle' },
        fill: {
            type: 'pattern',
            pattern: 'solid', //darkVertical
            fgColor: {
                argb: '0000FF',
            },
        },
    };
    const header = {
        border: true,
        money: false,
        height: 70,
        font: { size: 15, bold: false, color: { argb: 'FFFFFF' } },
        alignment: { horizontal: 'center', vertical: 'middle' },
        fill: {
            type: 'pattern',
            pattern: 'solid', //darkVertical
            fgColor: {
                argb: 'FF0000',
            },
        },
    };
    const data = {
        border: true,
        money: true,
        height: 0,
        font: { size: 12, bold: false, color: { argb: '000000' } },
        alignment: null,
        fill: null,
    };
    const footer = {
        border: true,
        money: true,
        height: 70,
        font: { size: 15, bold: true, color: { argb: 'FFFFFF' } },
        alignment: null,
        fill: {
            type: 'pattern',
            pattern: 'solid', //darkVertical
            fgColor: {
                argb: '0000FF',
            },
        },
    };
    if (widths && widths.length > 0) {
        ws.columns = widths;
    }

    let row = addRow(ws, [report], title);
    mergeCells(ws, row, 1, columns);

    addRow(ws, myHeader, header);
    // console.log('wb', wb);
    myData.forEach((row) => {
        addRow(ws, Object.values(row), data);
    });
    // console.log('myFooter', myFooter);

    row = addRow(ws, myFooter, footer);
    mergeCells(ws, row, 1, columns - 1);

    const buf = await wb.xlsx.writeBuffer();
    saveAs(new Blob([buf]), `${fileName}.xlsx`);
}

function addRow(ws, data, section) {
    const borderStyles = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
    };
    const row = ws.addRow(data);
    console.log('addRow', section, data);
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        if (section?.border) {
            cell.border = borderStyles;
        }
        if (section?.money && typeof cell.value === 'number') {
            cell.alignment = { horizontal: 'right', vertical: 'middle' };
            cell.numFmt = '$#,##0.00;[Red]-$#,##0.00';
        }
        if (section?.alignment) {
            cell.alignment = section.alignment;
        } else {
            cell.alignment = { vertical: 'middle' };
        }
        if (section?.font) {
            cell.font = section.font;
        }
        if (section?.fill) {
            cell.fill = section.fill;
        }
    });
    if (section?.height > 0) {
        row.height = section.height;
    }
    return row;
}

function mergeCells(ws, row, from, to) {
    // console.log(
    // 	'mergeCells',
    // 	row,
    // 	from,
    // 	to,
    // 	row.getCell(from)._address,
    // 	row.getCell(to)._address
    // );
    ws.mergeCells(`${row.getCell(from)._address}:${row.getCell(to)._address}`);
}

function columnToLetter(column) {
    var temp,
        letter = '';
    while (column > 0) {
        temp = (column - 1) % 26;
        letter = String.fromCharCode(temp + 65) + letter;
        column = (column - temp - 1) / 26;
    }
    return letter;
}
