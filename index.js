const XLSX = require('xlsx');

let readOptions = {
    type: 'string',
    cellFormula: false,
    cellText: false,
    cellDates: false
};

let workbookShipments = XLSX.readFile('./data/initial/Granngarden AB/Granngarden shipments 2019.xls', readOptions);
let shipmentsJSON = XLSX.utils.sheet_to_json(workbookShipments.Sheets['Sheet']);
let workbookInvoiceLines = XLSX.readFile('./data/initial/Granngarden AB/Granngarden lines 2019.xlsx', readOptions);
let linesJSON = XLSX.utils.sheet_to_json(workbookInvoiceLines.Sheets['Sheet']);
let reportJSON = [];
let lineTypes = [];
let currencies = [];

function addNewShipmentWithRevenue (shipment, invoiceLines) {
    checkAndAddLineTypes(invoiceLines);
    checkAndAddCurrencyTypes(invoiceLines);

    let newShipmentWithRevenue = {...shipment};
    newShipmentWithRevenue['Invoice Nr'] = invoiceLines[0]['Invoice Number'];

    for (currency of currencies) {
        newShipmentWithRevenue[`Total ${currency}`] = 0;
    }

    for (lineType of lineTypes) {
        let foundLineInvoiceByType = invoiceLines.find(l => l['Article Code'] === lineType['Article Code']);
        if (foundLineInvoiceByType == undefined) continue;
        let articleName = lineType['Article Name'];
        let netSum = foundLineInvoiceByType['Net Sum'];
        let currency = foundLineInvoiceByType['Currency'];
        newShipmentWithRevenue[articleName] = netSum;
        newShipmentWithRevenue[`${articleName} Currency`] = currency;
        newShipmentWithRevenue[`Total ${currency}`] += netSum;
    }

    if (newShipmentWithRevenue[`Total ${currency}`] === 0) delete newShipmentWithRevenue[`Total ${currency}`];
    
    reportJSON.push(newShipmentWithRevenue);
}

function checkAndAddLineTypes (linesInvoice) {
    for (line of linesInvoice) {
        if (lineTypes.findIndex(l => l['Article Code'] === line['Article Code']) === -1) {
            let articleName = '';
            if (line['Article Code'] === 'DTSPECIAL') {
                articleName = 'Extra costs'
            } else {
                articleName = line['Article Name'];
            }
            let newLineType = {
                'Article Code': line['Article Code'],
                'Article Name': articleName
            }
            lineTypes.push(newLineType);
        }
    }
}

function checkAndAddCurrencyTypes (linesInvoice) {
    for (line of linesInvoice) {
        if (!currencies.includes(line['Currency'])) {
            currencies.push(line['Currency']);
        }
    }
}

for (shipment of shipmentsJSON) {
    let foundLines = linesJSON.filter(line => line['Shipment ID'] === shipment['Shipment ID']);
    if (foundLines.length === 0) {
        console.log(`No invoice lines found for shipment ${shipment['Shipment ID']}!`);
        break;
    }

    addNewShipmentWithRevenue(shipment, foundLines);
}

let reportWorksheet = XLSX.utils.json_to_sheet(reportJSON, {
    cellDates: false
});

let newWorkbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(newWorkbook, reportWorksheet, 'data');

XLSX.writeFile(newWorkbook, './data/Granngarden AB 2019 Shipments Report.xlsx');