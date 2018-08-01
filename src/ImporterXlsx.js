import XLSX from 'xlsx'
import SpreadsheetColumn from 'spreadsheet-column'

export class ImporterXlsx {
  constructor() {
    // Stores the sheet by there name
    this.sheets = new Map()

    // store the sheet names in the right order
    this._sheetNames = []

    // The first column starts with '0'
    this.converter = new SpreadsheetColumn({ zero: true })
  }

  /**
   * Free some memory
   */
  clear() {
    this._sheetNames = []
    this.sheets = new Map()
  }

  /**
   * Opens the Spreadsheet and loads it
   * @param fileName {string} The file to open
   */
  async loadFile(fileName) {
    this.sheets = new Map()
    this._sheetNames = []
    const workbook = XLSX.readFile(fileName)

    // store the sheet in the internal map
    workbook.SheetNames.forEach(sheetName => {
      this._sheetNames.push(sheetName)
      const sheet = workbook.Sheets[sheetName]
      this.sheets.set(sheetName, sheet)
    })
  }

  /**
   * Returns the Cell value from a given sheet
   * @param sheetName {string} The name of the sheet
   * @param column {number} The column number start with '0'
   * @param row {number} The row number start with '0'
   * @return value {string} The Cell value
   */
  cellValue(sheetName, column, row) {
    const sheet = this.sheets.get(sheetName)
    if (sheet !== undefined) {
      const spreadRow = row + 1
      const spreadColumn = this.converter.fromInt(column)
      const address = `${spreadColumn}${spreadRow}`
      const cell = sheet[address]
      if (cell !== undefined) {
        if (typeof cell.v === 'string') {
          const val = cell.v.trim()
          if (val.length === 0) {
            return undefined
          }
          return val
        }
        return cell.v
      }
    }
  }

  /**
   * Returns a list of sheet names
   * @return sheets {array} A list of sheet names
   */
  get sheetNames() {
    return this._sheetNames
  }
}
