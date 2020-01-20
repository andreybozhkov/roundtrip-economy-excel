const XLSX = require('xlsx');

let readOptions = {
    type: 'string',
    cellFormula: false,
    cellText: false,
    cellDates: false
};

let workbookCustomerReferences = XLSX.readFile('./data/initial/Holmen Paper/2019-11 and 12/edge_protectors_2019_11_and_12.xlsx', readOptions);
let customerRefsJSON = XLSX.utils.sheet_to_json(workbookCustomerReferences.Sheets['edge_protectors']);
let workbookShipments = XLSX.readFile('./data/initial/Holmen Paper/2019-11 and 12/shipments.xls', readOptions);
let shipmentsJSON = XLSX.utils.sheet_to_json(workbookShipments.Sheets['shipments']);
let workbookProjects = XLSX.readFile('./data/initial/Holmen Paper/2019-11 and 12/projects.xls', readOptions);
let projectsJSON = XLSX.utils.sheet_to_json(workbookProjects.Sheets['projects']);
let reportJSON = [];

function addNewRoundTrip (customerRefLine, shipment) {
    let newRoundTrip = {...customerRefLine};
    newRoundTrip['Export ID'] = shipment.ID;
    newRoundTrip['Export Trailer'] = shipment.Trailer;
    newRoundTrip['Export Project Reporting Date'] = shipment['Project Reporting Date'];
    newRoundTrip['Export Start Date'] = shipment['Start Date'];
    newRoundTrip['Export End Date'] = shipment['End Date'];
    newRoundTrip['Export Traffic Line Group'] = shipment['Traffic Line Group'];
    newRoundTrip['Export Traffic Line'] = shipment['Traffic Line'];
    newRoundTrip['Export Estimated Net Profit'] = shipment['Estimated Net Profit'];
    newRoundTrip['Export Net Profit'] = shipment['Net Profit'];
    newRoundTrip['Export Customer Companies from Shipments'] = shipment['Customer Companies from Shipments'];
    reportJSON.push(newRoundTrip);
}

for (customerRefLine of customerRefsJSON) {
    for (shipment of shipmentsJSON) {
        if (shipment.Trailer === customerRefLine.Trailer && shipment['Start Date'] > customerRefLine['Start Date']) {
            addNewRoundTrip(customerRefLine, shipment);
            break;
        }
    }
}

let reportWorksheet = XLSX.utils.json_to_sheet(reportJSON, {
    cellDates: false
});

let newWorkbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(newWorkbook, reportWorksheet, 'data');

XLSX.writeFile(newWorkbook, './data/Holmen Paper report.xlsx');