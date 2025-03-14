[**@xhubiotable/importer-xlsx**](../README.md)

***

[@xhubiotable/importer-xlsx](../globals.md) / ImporterXlsx

# Class: ImporterXlsx

Defined in: [ImporterXlsx.ts:23](https://github.com/xhubioTable/importer-xlsx/blob/7a565c80f28a805aa445fdd2330eb66e31140b63/src/ImporterXlsx.ts#L23)

An importer for loading Excel spreadsheets.

This class implements the ImporterInterface to load Excel files,
store the sheets internally, and provide access to individual cell values.
It uses the XLSX library to read the file and SpreadsheetColumn to convert column indices.

## Implements

- [`ImporterInterface`](../interfaces/ImporterInterface.md)

## Constructors

### new ImporterXlsx()

> **new ImporterXlsx**(`opts`): [`ImporterXlsx`](ImporterXlsx.md)

Defined in: [ImporterXlsx.ts:48](https://github.com/xhubioTable/importer-xlsx/blob/7a565c80f28a805aa445fdd2330eb66e31140b63/src/ImporterXlsx.ts#L48)

Constructs a new ImporterXlsx instance.

#### Parameters

##### opts

[`ImporterXlsxOptions`](../interfaces/ImporterXlsxOptions.md) = `{}`

Options for initializing the importer, including an optional logger.

#### Returns

[`ImporterXlsx`](ImporterXlsx.md)

## Properties

### \_sheetNames

> **\_sheetNames**: `string`[] = `[]`

Defined in: [ImporterXlsx.ts:35](https://github.com/xhubioTable/importer-xlsx/blob/7a565c80f28a805aa445fdd2330eb66e31140b63/src/ImporterXlsx.ts#L35)

An array of sheet names in the order they appear in the workbook.

***

### converter

> **converter**: `SpreadsheetColumn`

Defined in: [ImporterXlsx.ts:41](https://github.com/xhubioTable/importer-xlsx/blob/7a565c80f28a805aa445fdd2330eb66e31140b63/src/ImporterXlsx.ts#L41)

A converter to map Excel column letters to numbers.
The conversion is configured such that the first column corresponds to 0.

***

### filename?

> `optional` **filename**: `string`

Defined in: [ImporterXlsx.ts:53](https://github.com/xhubioTable/importer-xlsx/blob/7a565c80f28a805aa445fdd2330eb66e31140b63/src/ImporterXlsx.ts#L53)

The name of the loaded file.

***

### logger

> **logger**: `LoggerInterface`

Defined in: [ImporterXlsx.ts:25](https://github.com/xhubioTable/importer-xlsx/blob/7a565c80f28a805aa445fdd2330eb66e31140b63/src/ImporterXlsx.ts#L25)

Logger instance used for logging operations.

#### Implementation of

[`ImporterInterface`](../interfaces/ImporterInterface.md).[`logger`](../interfaces/ImporterInterface.md#logger)

***

### sheets

> **sheets**: `Map`\<`string`, `WorkSheet`\>

Defined in: [ImporterXlsx.ts:30](https://github.com/xhubioTable/importer-xlsx/blob/7a565c80f28a805aa445fdd2330eb66e31140b63/src/ImporterXlsx.ts#L30)

A map storing the loaded worksheets by their sheet names.

## Accessors

### sheetNames

#### Get Signature

> **get** **sheetNames**(): `string`[]

Defined in: [ImporterXlsx.ts:157](https://github.com/xhubioTable/importer-xlsx/blob/7a565c80f28a805aa445fdd2330eb66e31140b63/src/ImporterXlsx.ts#L157)

Retrieves the sheet names of the loaded workbook in their original order.

##### Returns

`string`[]

An array of sheet names.

A list of all sheet names loaded from the file.

#### Implementation of

[`ImporterInterface`](../interfaces/ImporterInterface.md).[`sheetNames`](../interfaces/ImporterInterface.md#sheetnames)

## Methods

### cellValue()

> **cellValue**(`sheetName`, `column`, `row`): `undefined` \| `string` \| `number`

Defined in: [ImporterXlsx.ts:100](https://github.com/xhubioTable/importer-xlsx/blob/7a565c80f28a805aa445fdd2330eb66e31140b63/src/ImporterXlsx.ts#L100)

Retrieves the value of a cell from a specified sheet.

#### Parameters

##### sheetName

`string`

The name of the sheet from which to retrieve the cell value.

##### column

`number`

The column number (0-based).

##### row

`number`

The row number (0-based).

#### Returns

`undefined` \| `string` \| `number`

The cell value as a number, string, or undefined if the cell is empty.

#### Implementation of

[`ImporterInterface`](../interfaces/ImporterInterface.md).[`cellValue`](../interfaces/ImporterInterface.md#cellvalue)

***

### cellValueString()

> **cellValueString**(`sheetName`, `column`, `row`): `undefined` \| `string`

Defined in: [ImporterXlsx.ts:138](https://github.com/xhubioTable/importer-xlsx/blob/7a565c80f28a805aa445fdd2330eb66e31140b63/src/ImporterXlsx.ts#L138)

Retrieves the value of a cell from a specified sheet as a string.

#### Parameters

##### sheetName

`string`

The name of the sheet.

##### column

`number`

The column number (0-based).

##### row

`number`

The row number (0-based).

#### Returns

`undefined` \| `string`

The cell value as a string, or undefined if the cell is empty.

#### Implementation of

[`ImporterInterface`](../interfaces/ImporterInterface.md).[`cellValueString`](../interfaces/ImporterInterface.md#cellvaluestring)

***

### clear()

> **clear**(): `void`

Defined in: [ImporterXlsx.ts:60](https://github.com/xhubioTable/importer-xlsx/blob/7a565c80f28a805aa445fdd2330eb66e31140b63/src/ImporterXlsx.ts#L60)

Clears all loaded data to free up memory.

This method deletes any cached or loaded data from the importer.

#### Returns

`void`

#### Implementation of

[`ImporterInterface`](../interfaces/ImporterInterface.md).[`clear`](../interfaces/ImporterInterface.md#clear)

***

### loadFile()

> **loadFile**(`fileName`): `Promise`\<`void`\>

Defined in: [ImporterXlsx.ts:75](https://github.com/xhubioTable/importer-xlsx/blob/7a565c80f28a805aa445fdd2330eb66e31140b63/src/ImporterXlsx.ts#L75)

Opens and loads a file.

This method loads the specified file (which could be a spreadsheet or any other supported format)
and prepares its contents for further processing.

#### Parameters

##### fileName

`string`

The path or name of the Excel file to load.

#### Returns

`Promise`\<`void`\>

A promise that resolves when the file has been successfully loaded.

#### Implementation of

[`ImporterInterface`](../interfaces/ImporterInterface.md).[`loadFile`](../interfaces/ImporterInterface.md#loadfile)
