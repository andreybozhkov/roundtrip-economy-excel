const XLSX = require('xlsx');

let readOptions = {
    type: 'string',
    cellFormula: false,
    cellText: false,
    cellDates: false
};

let workbookShipments = XLSX.readFile('./data/initial/Zarges AB/shipments-2019-clean.xls', readOptions);
let shipmentsJSON = XLSX.utils.sheet_to_json(workbookShipments.Sheets['shipment_data']);
let workbookProjects = XLSX.readFile('./data/initial/Zarges AB/project-profit-export-clean.xls', readOptions);
let projectsJSON = XLSX.utils.sheet_to_json(workbookProjects.Sheets['project_profit_data']);
let reportJSON = [];

function addNewProjectProfit (shipment, project) {
    let newShipmentWithProjectProfit = {...shipment};
    newShipmentWithProjectProfit['Project Estimated Profit'] = project['Estimated Net Profit'];
    reportJSON.push(newShipmentWithProjectProfit);
}

for (shipment of shipmentsJSON) {
    for (project of projectsJSON) {
        if (shipment['Project ID'] === project['Project ID']) {
            addNewProjectProfit(shipment, project);
            break;
        }
    }
}

let reportWorksheet = XLSX.utils.json_to_sheet(reportJSON, {
    cellDates: false
});

let newWorkbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(newWorkbook, reportWorksheet, 'data');

XLSX.writeFile(newWorkbook, 'report.xlsx');