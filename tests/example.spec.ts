import { test } from "@playwright/test";
import * as Excel from "exceljs";
import { ExtractorResponse, Status } from "./utils";

interface HeaderValues {
  headerName?: string;
  headerValue?: string;
}

interface RowHeaders {
  conditionals: string[];
  headers: string[];
}

const getServiceType = (sheet: Excel.Worksheet, number: Number) => {
  const serviceName = sheet.name;
  if (serviceName) {
    return {
      status: Status.VALID,
      name: serviceName,
    };
  } else {
    return {
      status: Status.INVALID,
      message: "Sheet No " + number + ": " + "No Sheet Name found",
    };
  }
};

const getRowNumberOf = (sheet: Excel.Worksheet, columnName: string) => {
  let numberRow: number | undefined = undefined;
  sheet.eachRow({ includeEmpty: false }, function (row, number) {
    const lineitem = Array.isArray(row) ? row : row?.values;
    if (lineitem && Array.isArray(lineitem)) {
      lineitem.shift();
      if (lineitem.find((item) => columnName === item)) {
        numberRow = number;
      }
    }
  });
  return numberRow;
};

const validateSheet = (sheet: Excel.Worksheet, number: Number, sheeteName: string | undefined) => {
  if (sheeteName !== null) {
    const rowHeaderOrigin = getRowNumberOf(sheet, "Origin");
    const rowHeaderDestination = getRowNumberOf(sheet, "Destination");
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
              message: "Sheet No " + number + ": " + "No line items found",
            };
          }
        } else {
          return {
            message:
              "Sheet No " +
              number +
              ": " +
              "Header columns must no be empty keep the format",
          };
        }
      } else {
        return {
          message: "Sheet No " + number + ": " + "No Origin Header name found",
        };
      }
    } else {
      return {
        message:
          "Sheet No " +
          number +
          ": " +
          "No Origin and Destination Header name found",
      };
    }
  } else {
    return {
      message:
        "Sheet No " +
        sheeteName +
        ": " +
        "Invalid service name or invalid format",
    };
  }
};

const getTableHeaders = (sheet: Excel.Worksheet): HeaderValues[] => {
  const headerValues: HeaderValues[] = [];
  const tableHeaders: string[] = ["Service Name", "Charge Type", "UoM"];

  tableHeaders.forEach((name) => {
    const fieldNumber = getRowNumberOf(sheet, name);
    if (typeof fieldNumber === "number") {
      const row = sheet.getRow(fieldNumber);
      const lineitem = Array.isArray(row) ? row : row.values;
      if (lineitem && Array.isArray(lineitem)) {
        const cleanArray = lineitem.filter((item) => item);
        headerValues.push({
          headerName: cleanArray[0] ? cleanArray[0] : "",
          headerValue: cleanArray[1] ? cleanArray[1] : "",
        });
      }
    }
  });
  return headerValues;
};

const dynamicMapper = (rowType: string, serviceName: string) => {
  //Mandatory fields

  if (serviceName !== null) {
    return rowType == 'LANE' ? RateCardLineItemsHeaderLANEXPDD : RateCardLineItemsHeaderORADSTXPDD
  } else {
    return RateCardLineItemsHeaderLANEONFC
  }
}

const getRowType = (row: any, rowHeaders: RowHeaders) => {
  if (row) {
    //Row type puede ser mandatorio o puede ser la columna en azul que seria field
    if (row[0] === 'Conditional') {
      rowHeaders.conditionals.push(row)
    } else if (row[0] === 'Rating Code') {
      rowHeaders.headers.push(row)
    }
    return rowHeaders
  }
}

export const validateRow = (row: Excel.Row | any[], serviceName: string, number: Number, sheet: number) => {
  let validationError = []
  let cleanLineItem = []
  const lineitem = Array.isArray(row) ? row : row?.values

  if (lineitem && Array.isArray(lineitem)) {
    lineitem.shift()
    const headerValues: RowHeaders;
    const rowType = getRowType(lineitem, headerValues)
    if (lineitem.toString() !== null) {
      // const map = dynamicMapper(rowType, serviceName)
    }
  }
  return {
    status: validationError.length > 0 ? Status.INVALID : Status.VALID,
    data: validationError.length > 0 ? validationError : cleanLineItem,
  }
}

const extractor = (workbook: Excel.Workbook): ExtractorResponse => {
  let response: ExtractorResponse = {
    status: Status.VALID,
    validationErrors: [],
    data: [],
    numberRecordsProcessed: 0,
  };

  workbook.eachSheet((sheet, sheetNumber) => {
    const serviceType = getServiceType(sheet, sheetNumber);
    const rowHeaderOrigin = getTableHeaders(sheet);
    if (serviceType.status == Status.VALID) {
      const serviceName = serviceType.name;
      const isSheetInValid = validateSheet(sheet, sheetNumber, serviceName);
      if (Boolean(isSheetInValid)) {
        response.status = Status.INVALID;
      } else {
        if (sheet.actualRowCount > 4) {
          sheet.eachRow({ includeEmpty: false }, function (row, number) {
            if (number >= 5) {
              response.numberRecordsProcessed++
              const linteItem = row
              const hasInvalidRow = validateRow(linteItem, serviceName, number, sheetNumber)
            }
          });
        } else {
          response.status = Status.INVALID;
          const message =
            "Sheet No. " + sheetNumber + ": " + "No line item preset";
          response.validationErrors.push(message);
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

test("Read xlsx data info and store in objects", async ({ request }) => {
  const workbook = new Excel.Workbook();
  await workbook.xlsx.readFile("resources/Northline_StandardTemplate.xlsx");
  const response = extractor(workbook);
});

export default extractor;