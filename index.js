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

let errorLog = [];

function findShipment (customerRef) {
    let foundShipment = shipmentsJSON.find(shipment => shipment['Reference'] === customerRef);
    if (foundShipment === undefined) {
        let errorMsg = {
            'returnType': 'Error',
            'errorName': 'Shipment with this customer reference not found',
            'customerReference': customerRef
        };
        errorLog.push(errorMsg);
        return errorMsg;
    }
    else return foundShipment;
}

function findProject (shipmentProjectID) {
    let foundProject = projectsJSON.find(project => project.ID === shipmentProjectID);
    if (foundProject === undefined) {
        let errorMsg = {
            'returnType': 'Error',
            'errorName': 'Project with given ID not found',
            'projectID': shipmentProjectID
        };
        errorLog.push(errorMsg);
        return errorMsg;
    }
    else return foundProject;
}

function addNewEdgeProtectorLine (customerRefLine, shipment, project) {
    let newEdgeProtectorLine = {...customerRefLine};
    newEdgeProtectorLine['Project'] = project.ID;
    newEdgeProtectorLine['Haulier'] = shipment['Pickup Carrier Name'];
    newEdgeProtectorLine['Trailer nr'] = project['Trailer'];
    newEdgeProtectorLine['Invoice nr'] = '';
    newEdgeProtectorLine['Credit Note nr'] = '';
    newEdgeProtectorLine['Project Status'] = project['Status'];
    newEdgeProtectorLine['Replacement Project'] = 2044801;
    newEdgeProtectorLine['Notes'] = '';
    reportJSON.push(newEdgeProtectorLine);
}

for (customerRefLine of customerRefsJSON) {
    let foundShipment = findShipment(customerRefLine['Transport Booking']);
    if (foundShipment.returnType === 'Error') {
        console.log(foundShipment);
        continue;
    }

    let foundProject = findProject(foundShipment['Project ID'])
    if (foundProject.returnType === 'Error') {
        console.log(foundProject);
        continue;
    }
    
    addNewEdgeProtectorLine(customerRefLine, foundShipment, foundProject);
}

console.log(errorLog);

let reportWorksheet = XLSX.utils.json_to_sheet(reportJSON, {
    cellDates: false
});

let newWorkbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(newWorkbook, reportWorksheet, 'data');

XLSX.writeFile(newWorkbook, './data/Holmen Paper report.xlsx');