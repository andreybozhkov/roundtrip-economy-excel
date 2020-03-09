const XLSX = require('xlsx');

let readOptions = {
    type: 'string',
    cellFormula: false,
    cellText: false,
    cellDates: false
};

let workbookShipments = XLSX.readFile('./data/initial/Granngarden AB/Granngarden shipments 2019.xls', readOptions);
let shipmentsJSON = XLSX.utils.sheet_to_json(workbookShipments.Sheets['Sheet']);
let workbookInvoiceLines = XLSX.readFile('./data/initial/Granngarden AB/Granngarden lines 2019.xls', readOptions);
let linesJSON = XLSX.utils.sheet_to_json(workbookInvoiceLines.Sheets['Sheet']);
let reportJSON = [];

function addNewShipmentWithRevenue (shipment, lineInvoice) {
    let newShipmentWithProjectProfit = {...shipment};
    newShipmentWithProjectProfit['Project Estimated Profit'] = lineInvoice['Estimated Net Profit'];
    reportJSON.push(newShipmentWithProjectProfit);
}

for (shipment of shipmentsJSON) {
    for (lineInvoice of linesJSON) {
        if (shipment['ID'] === lineInvoice['Project ID']) {
            addNewProjectProfit(shipment, lineInvoice);
            break;
        }
    }
}

let reportWorksheet = XLSX.utils.json_to_sheet(reportJSON, {
    cellDates: false
});

let newWorkbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(newWorkbook, reportWorksheet, 'data');

XLSX.writeFile(newWorkbook, './data/Granngarden AB 2019 Shipments Reprot.xlsx');