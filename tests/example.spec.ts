import { test } from "@playwright/test";
import * as  Excel from "exceljs";


const RateCardLineItemsHeadeBasic = [
  { name: 'MIN', isrequired: false, type: 'number' },
  { name: 'BASIC', isrequired: false, type: 'number' },
]

const RateCardLineItemsHeaderLANEXPDD = [
  { name: 'LANE', isrequired: true, type: 'string' },
  ...RateCardLineItemsHeadeBasic,
  { name: '1-9999 KILO', isrequired: true, type: 'number' },
  { name: 'MAX KG', isrequired: false, type: 'number' },
  { name: 'Charge by', isrequired: false, type: 'string' },
]

const RateCardLineItemsHeaderORADSTXPDD = [
  { name: 'ORADST', isrequired: true, type: 'string' },
  ...RateCardLineItemsHeadeBasic,
  { name: '1-250 KILO Flat', isrequired: true, type: 'number' },
  { name: '250+ KILO Flat', isrequired: true, type: 'number' },
]

const RateCardLineItemsHeaderLANEONFC = [
  { name: 'LANE', isrequired: true, type: 'string' },
  { name: '3KG', isrequired: true, type: 'number' },
  { name: '5KG', isrequired: true, type: 'number' },
  { name: '>5KG BASIC', isrequired: true, type: 'number' },
  { name: '>5KG FLAT KILO', isrequired: true, type: 'number' },
]

const RateCardHeaderMap = {
  ORADST: { field: 'code' },
  LANE: { field: 'code' },
  BASIC: { field: 'baseCharge' },
  MIN: { field: 'minimumCharge' },
  'Charge by': { field: 'chargeBy' },
  'MAX KG': {
    field: 'maximumLimit',
    limit: 'weight',
  },
  '3KG': {
    line: {
      minimum: '0',
      maximum: '3',
      unit: 'KG',
      type: 'WEIGHT',
    },
  },
  '5KG': {
    line: {
      minimum: '3',
      maximum: '5',
      unit: 'KG',
      type: 'WEIGHT',
    },
  },
  '>5KG BASIC': {
    line: {
      minimum: '5',
      maximum: 'unlimited',
      unit: 'KG BASIC',
      type: 'WEIGHT',
    },
  },
  '>5KG FLAT KILO': {
    line: {
      minimum: '5',
      maximum: 'unlimited',
      unit: 'KG FLAT',
      type: 'WEIGHT',
    },
  },
  '1-9999 KILO': {
    line: {
      minimum: '1',
      maximum: '9999',
      unit: 'KG',
      type: 'WEIGHT',
    },
  },
  '1-250 KILO Flat': {
    line: {
      minimum: '1',
      maximum: '250',
      unit: 'KG FLAT',
      type: 'WEIGHT',
    },
  },
  '250+ KILO Flat': {
    line: {
      minimum: '251',
      maximum: 'unlimited',
      unit: 'KG FLAT',
      type: 'WEIGHT',
    },
  },
}

export interface ExtractorResponse {
  status: Status
  data: any
  validationErrors: string[]
  numberRecordsProcessed: number
}

export enum Status {
  VALID = 'valid',
  INVALID = 'invalid',
  SUCCESS = 'success',
  FAILURE = 'failure',
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

const getDimFactor = (sheet: Excel.Worksheet) => {
  const row = sheet.getRow(1)
  const lineitem = Array.isArray(row) ? row : row?.values
  if (lineitem && Array.isArray(lineitem)) {
    lineitem.shift() //remove first element which is null
    if (lineitem.length > 0) {
      const DimFactorArray = lineitem[0].split(' ')
      if (DimFactorArray.length >= 1) {
        if (DimFactorArray[0] === 'Section#1') {
          return DimFactorArray[DimFactorArray.length - 1]
        } else {
          return undefined
        }
      } else {
        return undefined
      }
    } else {
      return undefined
    }
  } else {
    return undefined
  }
}

const getRowNumberOf = (sheet: Excel.Worksheet, columnName: string) => {
  let numberRow: number | undefined
  sheet.eachRow({ includeEmpty: false }, function (row, number) {
    const lineitem = Array.isArray(row) ? row : row?.values
    if (lineitem && Array.isArray(lineitem)) {
      lineitem.shift()
      if (lineitem[0] === columnName) {
        numberRow = number
      }
    }
  })
  return numberRow
}

const validateSheet = (sheet: Excel.Worksheet, number: Number, serviceName: string | undefined) => {
  if (serviceName === 'Weight') {
    const rowHeaderLane = getRowNumberOf(sheet, 'Rating Code')
    const rowHeaderORADST = getRowNumberOf(sheet, 'Rating Code')
    if (typeof rowHeaderLane === 'number' && typeof rowHeaderORADST === 'number') {
      const rowLane = sheet.getRow(rowHeaderLane).values
      const rowORADST = sheet.getRow(rowHeaderORADST).values
      if (rowLane && Array.isArray(rowLane)) {
        if (rowORADST && Array.isArray(rowORADST)) {
          const rowLaneWithoutNull = rowLane.filter((rh) => rh !== null)
          const rowORADSTWithoutNull = rowORADST.filter((rh) => rh !== null)
          const cleanRowLaneWithoutNull = rowLaneWithoutNull.toString().replace(/(\r\n|\n|\r)/gm, "");
          const cleanrowRatingWithoutNull = rowORADSTWithoutNull.toString().replace(/(\r\n|\n|\r)/gm, "");
          if (cleanRowLaneWithoutNull === 'Rating Code,Origin,Destination,Basic Charge,Minimum Charge,0-50,51-100,101-500,501-1000,1001-3000,3000-more,Polys,200LT,1000LT,Dangerous GoodsRate,PUMP OUT EQUIPMENT FEE,DEMURRAGE,TAILGATE') {
            //match exactness
            const contentNumber = rowHeaderLane + 1
            const rowContents = sheet.getRow(contentNumber).values
            if (rowContents) {
              if (cleanrowRatingWithoutNull === 'Rating Code,Origin,Destination,Basic Charge,Minimum Charge,0-50,51-100,101-500,501-1000,1001-3000,3000-more,Polys,200LT,1000LT,Dangerous GoodsRate,PUMP OUT EQUIPMENT FEE,DEMURRAGE,TAILGATE') {
                //match exactness
                const contentNumber = rowHeaderLane + 1
                const rowContents = sheet.getRow(contentNumber).values
                if (rowContents) {
                  return null //structure is intact
                } else {
                  return {
                    message: 'Sheet No ' + number + ': ' + 'No lineitems found for ORADST',
                  }
                }
              } else {
                return {
                  message:
                    'Sheet No ' +
                    number +
                    ': ' +
                    'Header ORADST columns must be in same order and should have same key',
                }
              }
            } else {
              return {
                message: 'Sheet No ' + number + ': ' + 'No lineitems found for LANE',
              }
            }
          } else {
            return {
              message:
                'Sheet No ' + number + ': ' + 'Header  Lane columns must be in same order and should have same key',
            }
          }
        } else {
          return {
            message: 'Sheet No ' + number + ': ' + 'No ORADST Header name found',
          }
        }
      } else {
        return {
          message: 'Sheet No ' + number + ': ' + 'No LANE Header name found',
        }
      }
    } else {
      return {
        message: 'Sheet No ' + number + ': ' + 'No LANE and ORADST Header name found',
      }
    }
  } else {
    const rowHeaderLane = getRowNumberOf(sheet, 'LANE')
    if (typeof rowHeaderLane === 'number') {
      const rowLane = sheet.getRow(rowHeaderLane).values
      if (rowLane && Array.isArray(rowLane)) {
        const rowLaneWithoutNull = rowLane.filter((rh) => rh !== null)
        if (rowLaneWithoutNull.toString() === 'LANE,3KG,5KG,>5KG BASIC,>5KG FLAT KILO') {
          const contentNumber = rowHeaderLane + 1
          const rowContents = sheet.getRow(contentNumber).values
          if (rowContents) {
            return null //structure is intact
          } else {
            return {
              message: 'Sheet No ' + number + ': ' + 'No lineitems found for ORADST',
            }
          }
        } else {
          return {
            message:
              'Sheet No ' + number + ': ' + 'Header  Lane columns must be in same order and should have same key',
          }
        }
      } else {
        return {
          message: 'Sheet No ' + number + ': ' + 'No LANE Header name found',
        }
      }
    }
  }
}

const getRowType = (row: any, serviceName: string) => {
  if (serviceName === 'XPDD alltoall') {
    const rowType = row.length == 6 ? 'LANE' : 'ORADST'
    return rowType
  } else {
    return 'LANE'
  }
}

const provideMap = (rowType: string, serviceName: string) => {
  if (serviceName === 'XPDD alltoall') {
    return rowType == 'LANE' ? RateCardLineItemsHeaderLANEXPDD : RateCardLineItemsHeaderORADSTXPDD
  } else {
    return RateCardLineItemsHeaderLANEONFC
  }
}

const process = (data: any, dimfactor: string, serviceName: string | undefined) => {
  const rowType = getRowType(data, serviceName)
  const map = provideMap(rowType, serviceName)
  let lineitem = {
    code: '',
    lines: [],
    dimFactor: dimfactor,
    serviceName: serviceName,
    fromName: '',
    toName: '',
  }

  data.forEach((value, index) => {
    const object = RateCardHeaderMap[map[index].name]
    if (object.line) {
      // lineitem.lines.push({
      //   minimum: object.line.minimum,
      //   maximum: object.line.maximum,
      //   unit: object.line.unit,
      //   type: object.line.type,
      //   rate: value,
      // })
    } else {
      if (object.limit) {
        lineitem[object.field] = {}
        lineitem[object.field][object.limit] = value
      } else {
        lineitem[object.field] = value
      }
    }
  })
  return lineitem
}

export const validateRow = (row: Excel.Row | any[], serviceName: string | undefined, number: Number, sheet: number) => {
  let validationError = []
  let cleanLineItem = []
  const lineitem = Array.isArray(row) ? row : row?.values

  if (lineitem && Array.isArray(lineitem)) {
    lineitem.shift()
    const rowType = getRowType(lineitem, serviceName)
    if (lineitem[0] !== 'ORADST') {
      const map = provideMap(rowType, serviceName)
      const size = map.length
      if (lineitem.length === size) {
        for (let index = 0; index < lineitem.length; index++) {
          const value = lineitem[index]
          if (map[index].isrequired && (!value || value === '')) {
            validationError.push(
              // 'Sheet No ' + sheet + ': ' + 'Row No ' + number + ': Column "' + map[index].name + '" must Have Value'
            )
          }
          if (value) {
            // if (!checkType(value, map[index].type)) {
            //   validationError.push(
            //     'Sheet No ' +
            //     sheet +
            //     ': ' +
            //     'Row No ' +
            //     number +
            //     ': Column "' +
            //     map[index].name +
            //     '" must be of type ' +
            //     map[index].type
            //   )
            // }
          }
        }
        // cleanLineItem = lineitem
      } else {
        validationError.push(
          // 'Sheet No ' + sheet + ': ' + 'Row No ' + number + ': ' + 'Number of column is invalid, must be ' + size
        )
      }
    }
  }
  return {
    status: validationError.length > 0 ? Status.INVALID : Status.VALID,
    data: validationError.length > 0 ? validationError : cleanLineItem,
  }
}



test("Read xlsx data info and store in objects", async ({ request }) => {
  const workbook = new Excel.Workbook();
  await workbook.xlsx.readFile('resources/ef-37_template.xlsx')
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
    const dimfactor = getDimFactor(sheet)
    if (serviceType.status == Status.VALID) {
      const serviceName = serviceType.name

      const isSheetInvalid = validateSheet(sheet, sheetNumber, serviceName)
      if (Boolean(isSheetInvalid)) {
        response.status = Status.INVALID
        // response.validationErrors.push(isSheetInvalid.message)
      } else {
        if (sheet.actualRowCount > 4) {
          sheet.eachRow({ includeEmpty: false }, function (row, number) {
            if (number >= 5) {
              response.numberRecordsProcessed++
              const linteItem = row
              const hasInvalidRow = validateRow(linteItem, serviceName, number, sheetNumber)
              if (hasInvalidRow.status === Status.INVALID) {
                response.status = Status.INVALID
                response.validationErrors = [...response.validationErrors, ...hasInvalidRow.data]
              } else {
                const processedRow = process(hasInvalidRow.data, dimfactor, serviceName)
                response.data.push(processedRow)
              }
            }
          })
        } else {
          response.status = Status.INVALID
          const message = 'Sheet No. ' + sheetNumber + ': ' + 'No line item preset'
          response.validationErrors.push(message)
        }
      }
    } else {
      response.status = Status.INVALID;
      // response.validationErrors.push(serviceType.message)
    }
  })
  return response
}

export default extractor