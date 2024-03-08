import { test } from "@playwright/test";
import * as Excel from "exceljs";
import { ExtractorResponse, Status } from "./utils";
import * as R from 'ramda'

interface HeaderValues {
    headerName?: string;
    headerValue?: string;
}

interface RowHeaders {
    conditionals?: string[];
    headers?: string[];
}

interface CustomDynamicMap {
    name: string;
    isrequired: boolean;
    type: string;
}

const existCurrentSheet = (sheet: Excel.Worksheet, sheetName: string) => {
    if (sheetName) {
        return {
            status: Status.VALID,
            name: sheetName,
        };
    } else {
        return {
            status: Status.INVALID,
            message: "Sheet No " + sheetName + ": " + "No Sheet Name found",
        };
    }
};

const getTableHeaders = (sheet: Excel.Worksheet, sheetNumber: string): HeaderValues[] => {
    let validationError: string[] = []
    const headerValues: HeaderValues[] = [];
    const tableHeaders: string[] = ["Service Name", "Charge Type", "UoM"];

    tableHeaders.forEach((name) => {
        const worksheet = sheet.workbook.getWorksheet(sheetNumber)
        if (worksheet) {
            const fieldNumber = getRowNumberOf(sheet, name, sheetNumber);
            if (typeof fieldNumber === "number") {
                const row = worksheet?.getRow(fieldNumber)
                const lineitem = Array.isArray(row) ? row : row.values;
                if (lineitem && Array.isArray(lineitem)) {
                    const cleanArray = lineitem.filter((item) => item);
                    headerValues.push({
                        headerName: cleanArray[0] ? cleanArray[0] : "",
                        headerValue: cleanArray[1] ? cleanArray[1] : "",
                    });
                }
            }
        } else {
            validationError.push('Sheet No ' + sheet + ': ' + 'Row No ' + sheetNumber + ': ' + 'Number of column is invalid, must be ')
        }
    });
    return headerValues;
};

const getRowNumberOf = (sheet: Excel.Worksheet, columnName: string, sheetName: string) => {
    let numberRow: number | undefined = undefined;
    let validationError: string[] = []
    const worksheet = sheet.workbook.getWorksheet(sheetName)
    if (worksheet) {
        worksheet.eachRow({ includeEmpty: false }, function (row, number) {
            const lineitem = Array.isArray(row) ? row : row?.values;
            if (lineitem && Array.isArray(lineitem)) {
                lineitem.shift();
                if (lineitem.find((item) => columnName === item)) {
                    numberRow = number;
                }
            }
        });

    } else {
        validationError.push('Sheet No ' + sheet + ': ' + 'Row No ' + sheetName + ': ' + 'Number of column is invalid, must be ')
    }
    return numberRow;
};

const validateSheet = (sheet: Excel.Worksheet, sheetName: string) => {
    if (sheetName !== null) {
        const rowHeaderOrigin = getRowNumberOf(sheet, "Origin", sheetName);
        const rowHeaderDestination = getRowNumberOf(sheet, "Destination", sheetName);
        if (typeof rowHeaderOrigin === "number" && typeof rowHeaderDestination === "number") {
            const rowLane = sheet.getRow(rowHeaderOrigin).values;
            if (rowLane && Array.isArray(rowLane)) {
                const rowLaneWithoutNull = rowLane.filter((rh) => rh !== null);
                if (rowLaneWithoutNull.toString() !== null) {
                    const contentNumber = rowHeaderOrigin + 1;
                    const rowContents = sheet.getRow(contentNumber).values;
                    if (rowContents) {
                        return null;
                    } else {
                        return {
                            message: "Sheet No " + sheetName + ": " + "No line items found",
                        };
                    }
                } else {
                    return {
                        message: "Sheet No " + sheetName + ": " + "Header columns must no be empty keep the format",
                    };
                }
            } else {
                return {
                    message: "Sheet No " + sheetName + ": " + "No Origin Header name found",
                };
            }
        } else {
            return {
                message: "Sheet No " + sheetName + ": " + "No Origin and Destination Header name found",
            };
        }
    } else {
        return {
            message: "Sheet No " + sheetName + ": " + "Invalid service name or invalid format",
        };
    }
};

const getSheetRowHeaders = (sheet: Excel.Worksheet, sheetNumber: string, propertyName: string) => {
    let validationError: string[] = []
    let headerValues: any[] = []

    const worksheet = sheet.workbook.getWorksheet(sheetNumber)
    if (worksheet) {
        const fieldNumber = getRowNumberOf(sheet, propertyName, sheetNumber);
        if (typeof fieldNumber === "number") {
            const row = worksheet.getRow(fieldNumber);
            const lineitem = Array.isArray(row) ? row : row.values;
            if (lineitem && Array.isArray(lineitem)) {
                const cleanArray: any[string] = lineitem.filter((item) => item);
                cleanArray.forEach(element => {
                    headerValues?.push(element);
                });
            }
        }
    } else {
        validationError.push('Sheet No ' + sheet + ': ' + 'Row No ' + sheetNumber + ': ' + 'Number of column is invalid, must be ')
    }
    return headerValues;
};

const dynamicMap = (conditionals: string[], properties: string[]) => {
    const combineHeaders = R.zip(properties, conditionals)
    const customDymanicMap = R.map(([name, isrequired]) => ({
        name,
        isrequired: isrequired === 'Mandatory' ? true : false,
        type: name === 'Rating Code' || name === 'Origin' || name === 'Destination' ? 'string' : 'number'
    }), combineHeaders);
    return customDymanicMap;
};

const checkType = (value, type) => {
    return typeof value === type
}

const validateCurrentRow = (row: Excel.Row | any[], number: Number, sheet: string, map: any) => {
    let validationError: any = []
    let cleanLineItem: any[] = []
    const lineitem = Array.isArray(row) ? row : row?.values

    if (lineitem && Array.isArray(lineitem)) {
        lineitem.shift()
        const size = map.length
        if (lineitem.length === size) {
            for (let index = 0; index < lineitem.length; index++) {
                const value = lineitem[index]
                if (map[index].isrequired && (!value || value === '')) {
                    validationError.push('Sheet No ' + sheet + ': ' + 'Row No ' + number + ': Column "' + map[index].name + '" must Have Value')
                }
                if (value) {
                    if (!checkType(value, map[index].type)) {
                        validationError.push('Sheet No ' + sheet + ': ' + 'Row No ' + number + ': Column "' + map[index].name + '" must be of type ' + map[index].type)
                    }
                }
            }
            cleanLineItem = lineitem
        } else {
            validationError.push('Sheet No ' + sheet + ': ' + 'Row No ' + number + ': ' + 'Number of column is invalid, must be ' + size)
        }
    }
    return {
        status: validationError.length > 0 ? Status.INVALID : Status.VALID,
        data: validationError.length > 0 ? validationError : cleanLineItem,
    }
}

const process = (data: any, sheetName: string, tableHeaders: HeaderValues[], map: any) => {
    let lineitem = {
        code: '',
        lines: [],
        // dimFactor: dimfactor,
        serviceName: sheetName,
        fromName: '',
        toName: '',
    }
    //Build a dynamic obj first removing standart fields
    //in map[4] can build custom fields
    const object = {};
    data.forEach((value, index) => {
        const customValues = map[index];
        const element = map[index];
        if (index > 4) {
            object[element.name] = {
                line: {
                    minimum: data[4],
                    maximum: data[5],
                    unit: tableHeaders[2].headerValue,
                    type: tableHeaders[1].headerValue,
                }
            }
        }
        if (object) {
            lineitem.lines.push({ object })
        }
    })
    return lineitem
}

const extractor = (workbook: Excel.Workbook): ExtractorResponse => {
    let response: ExtractorResponse = {
        status: Status.VALID,
        validationErrors: [],
        data: [],
        numberRecordsProcessed: 0,
    };

    workbook.eachSheet((sheet, sheetNumber) => {
        const sheetName = sheet.name;
        const serviceType = existCurrentSheet(sheet, sheetName);
        if (serviceType.status == Status.VALID) {
            const isSheetInValid = validateSheet(sheet, sheetName);
            if (Boolean(isSheetInValid)) {
                response.status = Status.INVALID;
            } else {
                //Header at the top of the table
                const tableHeaders = getTableHeaders(sheet, sheetName);
                const conditinals = getSheetRowHeaders(sheet, sheetName, 'Conditional')
                const properties = getSheetRowHeaders(sheet, sheetName, 'Rating Code')
                if (conditinals && properties !== null) {
                    //Build dynamic map
                    const map = dynamicMap(conditinals, properties)
                    if (sheet.actualRowCount > 4) {
                        sheet.eachRow({ includeEmpty: false }, function (row, number) {
                            if (number >= 7) {
                                response.numberRecordsProcessed++
                                const linteItem = row
                                const hasInvalidRow = validateCurrentRow(linteItem, number, sheetName, map)
                                if (hasInvalidRow.status === Status.INVALID) {
                                    response.status = Status.INVALID
                                    response.validationErrors = [...response.validationErrors, ...hasInvalidRow.data]
                                } else {
                                    const processedRow = process(hasInvalidRow.data, sheetName, tableHeaders, map)
                                    response.data.push(processedRow)
                                }
                            }
                        })
                    } else {
                        response.status = Status.INVALID;
                        const message = "Sheet No. " + sheetNumber + ": " + "No line item preset";
                        response.validationErrors.push(message);
                    }
                }
            }
        } else {
            response.status = Status.INVALID;
            //validate this after
            // response.validationErrors.push(serviceType.message)
        }
    });
    return response;
};

export default extractor;

test("Read xlsx data info and store in objects", async ({ request }) => {
    const workbook = new Excel.Workbook();
    await workbook.xlsx.readFile("resources/Northline_StandardTemplate.xlsx");
    const response = extractor(workbook);
});