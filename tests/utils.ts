export enum Status {
    VALID = 'valid',
    INVALID = 'invalid',
    SUCCESS = 'success',
    FAILURE = 'failure',
  }

export interface ExtractorResponse {
    status: Status
    data: any
    validationErrors: string[]
    numberRecordsProcessed: number
  }