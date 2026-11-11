---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Import-OfficeExcel
## SYNOPSIS
Reads rows from a worksheet, range or table as objects.

## SYNTAX
### Document (Default)
```powershell
Import-OfficeExcel [-Sheet <OfficeSheetSelector[]>] [-NoHeader] [-As <ExcelImportOutput>] [-NumericAsDecimal] [-FormulaMode <OfficeExcelFormulaMode>] [-Culture <cultureinfo>] [-SchemaSampleRows <int>] [-ChunkRows <int>] [-Range <string>] [-Document <ExcelDocument>] [<CommonParameters>]
```

### Path
```powershell
Import-OfficeExcel [-Path] <string[]> [-Sheet <OfficeSheetSelector[]>] [-NoHeader] [-As <ExcelImportOutput>] [-NumericAsDecimal] [-FormulaMode <OfficeExcelFormulaMode>] [-Culture <cultureinfo>] [-SchemaSampleRows <int>] [-ChunkRows <int>] [-Range <string>] [<CommonParameters>]
```

### PathTable
```powershell
Import-OfficeExcel [-Path] <string[]> -Table <string> [-NoHeader] [-As <ExcelImportOutput>] [-NumericAsDecimal] [-FormulaMode <OfficeExcelFormulaMode>] [-Culture <cultureinfo>] [-SchemaSampleRows <int>] [-ChunkRows <int>] [<CommonParameters>]
```

### LiteralPath
```powershell
Import-OfficeExcel -LiteralPath <string[]> [-Sheet <OfficeSheetSelector[]>] [-NoHeader] [-As <ExcelImportOutput>] [-NumericAsDecimal] [-FormulaMode <OfficeExcelFormulaMode>] [-Culture <cultureinfo>] [-SchemaSampleRows <int>] [-ChunkRows <int>] [-Range <string>] [<CommonParameters>]
```

### Uri
```powershell
Import-OfficeExcel [-Uri] <uri> [-Sheet <OfficeSheetSelector[]>] [-NoHeader] [-As <ExcelImportOutput>] [-NumericAsDecimal] [-FormulaMode <OfficeExcelFormulaMode>] [-Culture <cultureinfo>] [-SchemaSampleRows <int>] [-ChunkRows <int>] [-Range <string>] [-AllowHttp] [<CommonParameters>]
```

### LiteralPathTable
```powershell
Import-OfficeExcel -Table <string> -LiteralPath <string[]> [-NoHeader] [-As <ExcelImportOutput>] [-NumericAsDecimal] [-FormulaMode <OfficeExcelFormulaMode>] [-Culture <cultureinfo>] [-SchemaSampleRows <int>] [-ChunkRows <int>] [<CommonParameters>]
```

### UriTable
```powershell
Import-OfficeExcel [-Uri] <uri> -Table <string> [-NoHeader] [-As <ExcelImportOutput>] [-NumericAsDecimal] [-FormulaMode <OfficeExcelFormulaMode>] [-Culture <cultureinfo>] [-SchemaSampleRows <int>] [-ChunkRows <int>] [-AllowHttp] [<CommonParameters>]
```

### DocumentTable
```powershell
Import-OfficeExcel -Table <string> [-NoHeader] [-As <ExcelImportOutput>] [-NumericAsDecimal] [-FormulaMode <OfficeExcelFormulaMode>] [-Culture <cultureinfo>] [-SchemaSampleRows <int>] [-ChunkRows <int>] [-Document <ExcelDocument>] [<CommonParameters>]
```

## DESCRIPTION
Reads rows from a worksheet, range or table as objects.

## EXAMPLES

### EXAMPLE 1
```powershell
Import-OfficeExcel -Path .\Reports\Q3-Sales.xlsx
```

Reads the used range of the first worksheet, using the header row as property names.

### EXAMPLE 2
```powershell
Import-OfficeExcel -Path .\Reports\Q3-Sales.xlsx -Sheet Orders -Range A5 |
                Where-Object Status -eq 'Open' |
                Sort-Object Amount -Descending |
                Select-Object -First 10 OrderId, Customer, Region, Amount
```

Reads the Orders sheet from row 5, where the header starts below a title block, and lists the ten largest open orders.

### EXAMPLE 3
```powershell
Import-OfficeExcel -Path .\Reports\Helpdesk.xlsx -Table Tickets |
                Where-Object Status -ne 'Closed' |
                Group-Object Assignee |
                Sort-Object Count -Descending |
                Select-Object Name, Count
```

Reads the rows of the Tickets table wherever it sits in the workbook and counts open tickets per assignee.

### EXAMPLE 4
```powershell
Get-ChildItem -Path .\Branches -Filter *.xlsx |
                Import-OfficeExcel -Sheet * |
                Export-OfficeExcel -Path .\Reports\All-Branches.xlsx -Sheet 'All branches' -Title 'Branch sales, October 2026'
```

Reads every worksheet of each branch workbook (each row carries its WorksheetName) and writes one consolidated workbook.

## PARAMETERS

### -AllowHttp
Allow HTTP workbook downloads in addition to HTTPS.

```yaml
Type: SwitchParameter
Parameter Sets: Uri, UriTable
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -As
Selects objects, hashtables, a DataTable, an owned IDataReader or column output.

```yaml
Type: ExcelImportOutput
Parameter Sets: Document, Path, PathTable, LiteralPath, Uri, LiteralPathTable, UriTable, DocumentTable
Aliases: None
Possible values: Object, Hashtable, DataTable, DataReader, Column

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ChunkRows
Worksheet row count requested from each streaming chunk when -AsDataReader is used.

```yaml
Type: Int32
Parameter Sets: Document, Path, PathTable, LiteralPath, Uri, LiteralPathTable, UriTable, DocumentTable
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Culture
Culture used for value conversion and number-format presets.

```yaml
Type: CultureInfo
Parameter Sets: Document, Path, PathTable, LiteralPath, Uri, LiteralPathTable, UriTable, DocumentTable
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Caller-owned workbook; omission uses the enclosing Excel content block.

```yaml
Type: ExcelDocument
Parameter Sets: Document, DocumentTable
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -FormulaMode
Formula read mode. CachedValue returns workbook cached results; FormulaText returns formula expressions when present.

```yaml
Type: OfficeExcelFormulaMode
Parameter Sets: Document, Path, PathTable, LiteralPath, Uri, LiteralPathTable, UriTable, DocumentTable
Aliases: None
Possible values: CachedValue, FormulaText

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LiteralPath
Literal workbook paths, including FileInfo.FullName and provider PSPath.

```yaml
Type: String[]
Parameter Sets: LiteralPath, LiteralPathTable
Aliases: PSPath, FullName
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -NoHeader
Writes data without a header row.

```yaml
Type: SwitchParameter
Parameter Sets: Document, Path, PathTable, LiteralPath, Uri, LiteralPathTable, UriTable, DocumentTable
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NumericAsDecimal
Prefer decimals instead of doubles for numeric values.

```yaml
Type: SwitchParameter
Parameter Sets: Document, Path, PathTable, LiteralPath, Uri, LiteralPathTable, UriTable, DocumentTable
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Workbook paths resolved with filesystem wildcard support.

```yaml
Type: String[]
Parameter Sets: Path, PathTable
Aliases: FilePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Range
A1 cell range on the selected worksheet.

```yaml
Type: String
Parameter Sets: Document, Path, LiteralPath, Uri
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SchemaSampleRows
Maximum row count inspected when -AsDataReader infers the reader schema.

```yaml
Type: Int32
Parameter Sets: Document, Path, PathTable, LiteralPath, Uri, LiteralPathTable, UriTable, DocumentTable
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Sheet
Worksheet name, one-based position or native worksheet object. Multi-sheet commands also accept wildcard names.

```yaml
Type: OfficeSheetSelector[]
Parameter Sets: Document, Path, LiteralPath, Uri
Aliases: WorksheetName
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Table
Native Excel table used to resolve the source range or header.

```yaml
Type: String
Parameter Sets: PathTable, LiteralPathTable, UriTable, DocumentTable
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Uri
Remote workbook URI to import.

```yaml
Type: Uri
Parameter Sets: Uri, UriTable
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.String[]`
- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `None`

## RELATED LINKS

- None
