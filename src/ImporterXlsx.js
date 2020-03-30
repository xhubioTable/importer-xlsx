import XLSX from 'xlsx'
import SpreadsheetColumn from 'spreadsheet-column'

/**
 * An importer to load Excel Spreadsheets. This importer is used by the file
 * processor. So it is possible to create importer for different kind of files.
 */
export class ImporterXlsx {
  constructor() {
    /** {Map} Stores the sheets by there name */
    this.sheets = new Map()

    // Stores the sheet names in the right order
    this._sheetNames = []

    /** A converter to map Excel columns to numbers. The first column starts with '0' */
    this.converter = new SpreadsheetColumn({ zero: true })
  }

  /**
   * Frees some memory. Cleats the existing loaded sheets
   */
  clear() {
    this._sheetNames = []
    this.sheets = new Map()
  }

  /**
   * Opens the Spreadsheet file and loads it.
   * @param fileName {string} The file to open
   */
  async loadFile(fileName) {
    this.sheets = new Map()
    this._sheetNames = []
    const workbook = XLSX.readFile(fileName)

    // store the sheet in the internal map
    workbook.SheetNames.forEach((sheetName) => {
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
   * Returns the list with the names of the loaded sheet in the original order
   * @return sheets {array} A list of sheet names
   */
  get sheetNames() {
    return this._sheetNames
  }
}
