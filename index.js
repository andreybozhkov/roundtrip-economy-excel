const XLSX = require('xlsx');

let readOptions = {
    type: 'string',
    cellFormula: false,
    cellText: false,
    cellDates: false
};

let workBookShipments = XLSX.readFile('./data/initial/NTG Multimodal costs/shipments_Italy_to_Sweden.xls', readOptions);
let shipmentsJSON = XLSX.utils.sheet_to_json(workBookShipments.Sheets['shipments']);
let workBookProjectsEstimates = XLSX.readFile('./data/initial/NTG Multimodal costs/estimates_projects.xls', readOptions);
let projectsEstimatesJSON = XLSX.utils.sheet_to_json(workBookProjectsEstimates.Sheets['projects_estimates']);
let reportJSON = [];
let projectsList = [];

function generateProjectsList (shipments) {
    for (shipment of shipments) {
        if (!projectsList.includes(shipment['Project ID'])) {
            projectsList.push(shipment['Project ID']);
        }
    }
}

function addNewShipment (shipment, project, firstShipment) {
    let newShipment = {...shipment};
    if (firstShipment) {
        newShipment['Project Estimated Cost'] = project['Project Estimated Cost'];
    }
    else {
        newShipment['Project Estimated Cost'] = 0;
    }
    newShipment['Project Estimated Cost Currency'] = project['Project Estimated Cost Currency'];
    reportJSON.push(newShipment);
}

generateProjectsList(shipmentsJSON);

for (uniqueProject of projectsList) {
    let firstShipment = true;
    for (shipment of shipmentsJSON) {
        if (shipment['Project ID'] === uniqueProject) {
            let project = projectsEstimatesJSON.find(project => project['Project ID'] === uniqueProject);
            addNewShipment(shipment, project, firstShipment);
            firstShipment = false;
        }
    }
}

let reportWorksheet = XLSX.utils.json_to_sheet(reportJSON, {
   cellDates: false
});

let newWorkbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(newWorkbook, reportWorksheet, 'data');

XLSX.writeFile(newWorkbook, './data/NTG Multimodal prices Italy to Sweden.xlsx');