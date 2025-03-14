[**@xhubiotable/importer-xlsx**](../README.md)

***

[@xhubiotable/importer-xlsx](../globals.md) / ImporterInterface

# Interface: ImporterInterface

Defined in: [ImporterInterface.ts:9](https://github.com/xhubioTable/importer-xlsx/blob/7a565c80f28a805aa445fdd2330eb66e31140b63/src/ImporterInterface.ts#L9)

Interface for an importer that loads spreadsheet files.

An importer is responsible for opening a file, reading its contents,
and providing access to the sheet names and cell values.

## Properties

### cellValue()

> **cellValue**: (`sheetName`, `columnNumber`, `rowNumber`) => `undefined` \| `string` \| `number`

Defined in: [ImporterInterface.ts:39](https://github.com/xhubioTable/importer-xlsx/blob/7a565c80f28a805aa445fdd2330eb66e31140b63/src/ImporterInterface.ts#L39)

Retrieves the value of a cell from a specified sheet.

#### Parameters

##### sheetName

`string`

The name of the sheet.

##### columnNumber

`number`

The column number, starting at 0.

##### rowNumber

`number`

The row number, starting at 0.

#### Returns

`undefined` \| `string` \| `number`

The value of the cell as a number, string, or undefined if the cell is empty.

***

### cellValueString()

> **cellValueString**: (`sheetName`, `columnNumber`, `rowNumber`) => `undefined` \| `string`

Defined in: [ImporterInterface.ts:53](https://github.com/xhubioTable/importer-xlsx/blob/7a565c80f28a805aa445fdd2330eb66e31140b63/src/ImporterInterface.ts#L53)

Retrieves the value of a cell from a specified sheet as a string.

#### Parameters

##### sheetName

`string`

The name of the sheet.

##### columnNumber

`number`

The column number, starting at 0.

##### rowNumber

`number`

The row number, starting at 0.

#### Returns

`undefined` \| `string`

The cell value as a string, or undefined if the cell is empty.

***

### clear()

> **clear**: () => `void`

Defined in: [ImporterInterface.ts:64](https://github.com/xhubioTable/importer-xlsx/blob/7a565c80f28a805aa445fdd2330eb66e31140b63/src/ImporterInterface.ts#L64)

Clears all loaded data to free up memory.

This method deletes any cached or loaded data from the importer.

#### Returns

`void`

***

### loadFile()

> **loadFile**: (`fileName`) => `Promise`\<`void`\>

Defined in: [ImporterInterface.ts:24](https://github.com/xhubioTable/importer-xlsx/blob/7a565c80f28a805aa445fdd2330eb66e31140b63/src/ImporterInterface.ts#L24)

Opens and loads a file.

This method loads the specified file (which could be a spreadsheet or any other supported format)
and prepares its contents for further processing.

#### Parameters

##### fileName

`string`

The path or name of the file to open.

#### Returns

`Promise`\<`void`\>

A promise that resolves when the file has been successfully loaded.

***

### logger

> **logger**: `LoggerInterface`

Defined in: [ImporterInterface.ts:13](https://github.com/xhubioTable/importer-xlsx/blob/7a565c80f28a805aa445fdd2330eb66e31140b63/src/ImporterInterface.ts#L13)

Logger instance for logging importer operations.

***

### sheetNames

> **sheetNames**: `string`[]

Defined in: [ImporterInterface.ts:29](https://github.com/xhubioTable/importer-xlsx/blob/7a565c80f28a805aa445fdd2330eb66e31140b63/src/ImporterInterface.ts#L29)

A list of all sheet names loaded from the file.
