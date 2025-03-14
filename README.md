**@xhubiotable/importer-xlsx**

***

# ImporterXlsx Documentation

This document describes the ImporterXlsx, an importer implementation
used to load Excel spreadsheets. The ImporterXlsx reads Excel files
using the XLSX library and provides methods to access sheet names and
individual cell values.

## Overview

ImporterXlsx is designed to: \* Load an Excel file and extract its
worksheets. \* Store the sheets internally using a Map for fast access.
\* Provide the original order of sheet names. \* Convert Excel column
letters to 0-based column numbers using the SpreadsheetColumn library.
\* Retrieve cell values as numbers or strings. \* Clear loaded data to
free memory.

This importer is typically used by the file processor to read and parse
spreadsheet data before it is transformed into a decision table or
similar internal model.

## Key Features

-   **File Loading:** Opens an Excel file and stores its sheets
    internally.

-   **Sheet Management:** Maintains an ordered list of sheet names and a
    mapping of sheet names to worksheet objects.

-   **Cell Access:** Retrieves cell values by converting 0-based column
    and row indices to Excel’s cell addresses. Provides methods to
    return cell values either in their native type or as strings.

-   **Data Clearance:** Supports clearing of loaded data to manage
    memory usage.

## Methods

    async loadFile(fileName: string): Promise<void>

Loads the Excel file specified by `fileName` using the XLSX library. The
method reads all sheets in the workbook, storing each sheet in an
internal map, and records the sheet names in their original order.

    cellValue(sheetName: string, column: number, row: number): number | string | undefined

Returns the value of a cell from the specified sheet. The column number
is converted to an Excel column letter and the row number is adjusted
from 0-based to 1-based indexing. If the cell value is a string, it is
trimmed; if empty, undefined is returned.

    cellValueString(sheetName: string, column: number, row: number): string | undefined

Retrieves the value of a cell as a string. If the cell value is numeric,
it is converted to a string.

    clear(): void

Clears the internal cache of loaded sheets and sheet names, freeing
memory.

    === Usage

    1. **Instantiate the Importer:**
       ```bash
       const importer = new ImporterXlsx({ logger: myLogger })
