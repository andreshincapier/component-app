import { test } from "@playwright/test";
import * as  Excel from "exceljs";
import { ExtractorResponse, Status } from './utils'

interface HeaderValues {
  headerName?: string;
  headerValue?: string;
}

const getServiceType = (sheet: Excel.Worksheet, number: Number) => {
  const serviceName = sheet.name
  if (serviceName) {
    return {
      status: Status.VALID,
      name: serviceName,
    }
  } else {
    return {
      status: Status.INVALID,
      message: 'Sheet No ' + number + ': ' + 'No Sheet Name found',
    }
  }
}

const getRowNumberOf = (sheet: Excel.Worksheet, columnName: string) => {
  let numberRow: number | undefined = undefined
  sheet.eachRow({ includeEmpty: false }, function (row, number) {
    const lineitem = Array.isArray(row) ? row : row?.values
    if (lineitem && Array.isArray(lineitem)) {
      lineitem.shift()
      if (lineitem.find((item) => columnName === item)) {
        numberRow = number
      }
    }
  })
  return numberRow
}

const validateSheet = (sheet: Excel.Worksheet, number: Number, sheeteName: string | undefined) => {
  if (sheeteName !== null) {
    const rowHeaderOrigin = getRowNumberOf(sheet, 'Origin')
    const rowHeaderDestination = getRowNumberOf(sheet, 'Destination')
    if (typeof rowHeaderOrigin === 'number' && typeof rowHeaderDestination === 'number') {
      const rowLane = sheet.getRow(rowHeaderOrigin).values
      if (rowLane && Array.isArray(rowLane)) {
        const rowLaneWithoutNull = rowLane.filter((rh) => rh !== null)
        if (rowLaneWithoutNull.toString() !== null) {
          const contentNumber = rowHeaderOrigin + 1
          const rowContents = sheet.getRow(contentNumber).values
          if (rowContents) {
            return null
          } else {
            return { message: 'Sheet No ' + number + ': ' + 'No line items found', }
          }
        } else {
          return { message: 'Sheet No ' + number + ': ' + 'Header columns must no be empty keep the format', }
        }
      } else {
        return { message: 'Sheet No ' + number + ': ' + 'No Origin Header name found', }
      }
    } else {
      return { message: 'Sheet No ' + number + ': ' + 'No Origin and Destination Header name found', }
    }
  } else {
    return { message: 'Sheet No ' + sheeteName + ': ' + 'Invalid service name or invalid format', }
  }
}

const getTableHeaders = (sheet: Excel.Worksheet): HeaderValues => {
  const tableHeaders: string[] = ['Service Name', 'Charge Type', 'UoM'];
  const herders = tableHeaders.map((headerName, index) => {
    const fieldRowNumber = getRowNumberOf(sheet, headerName)
    if (typeof fieldRowNumber === 'number') {
      const rowLane = sheet.getRow(fieldRowNumber).values
      return {
        headerName: rowLane[4] ? rowLane[4] : '',
        headerValue: rowLane[5] ? rowLane[5] : '',
      };
    }
  });
  return herders;
}

test("Read xlsx data info and store in objects", async ({ request }) => {
  const workbook = new Excel.Workbook();
  await workbook.xlsx.readFile('resources/Northline_StandardTemplate.xlsx')
  const response = extractor(workbook)
});

const extractor = (workbook: Excel.Workbook): ExtractorResponse => {
  let response: ExtractorResponse = {
    status: Status.VALID,
    validationErrors: [],
    data: [],
    numberRecordsProcessed: 0,
  }

  workbook.eachSheet((sheet, sheetNumber) => {
    const serviceType = getServiceType(sheet, sheetNumber)
    const rowHeaderOrigin = getTableHeaders(sheet)
    if (serviceType.status == Status.VALID) {
      const serviceName = serviceType.name
      const isSheetInValid = validateSheet(sheet, sheetNumber, serviceName)
      if (Boolean(isSheetInValid)) {
        response.status = Status.INVALID
      } else {
        if (sheet.actualRowCount > 4) {
          sheet.eachRow({ includeEmpty: false }, function (row, number) { })
        } else {
          response.status = Status.INVALID
          const message = 'Sheet No. ' + sheetNumber + ': ' + 'No line item preset'
          response.validationErrors.push(message)
        }
      }
    }
    else {
      response.status = Status.INVALID
      //validate this after
      // response.validationErrors.push(serviceType.message)
    }
  })
  return response
}

export default extractor