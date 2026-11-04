---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Export-OfficeExcel
## SYNOPSIS
Exports PowerShell objects to an Excel workbook using an operator-friendly surface.

## SYNTAX
### Create (Default)
```powershell
Export-OfficeExcel [-Path] <string> [-InputObject <Object>] [-Sheet <OfficeSheetSelector>] [-TableName <string>] [-Title <string>] [-Address <string>] [-Theme <OfficeThemeInput>] [-TableStyle <ExcelTableStyle>] [-NoTable] [-NoTableStyle] [-NoHeader] [-NoAutoFilter] [-NoAutoFit] [-NoFreeze] [-FreezeFirstColumn] [-ShowFirstColumn] [-ShowLastColumn] [-ShowColumnStripes] [-NoRowStripes] [-ColumnFormat <hashtable>] [-Decimals <int>] [-Culture <cultureinfo>] [-IgnoreMissingColumnFormat] [-ExcludeProperty <string[]>] [-CollectionSeparator <string>] [-FlattenOptions <OfficeObjectFlattenOptions>] [-DocumentProperty <hashtable>] [-DateSystem <ExcelDateSystem>] [-SaveOptions <ExcelSaveOptions>] [-Show] [-PassThru] [-NoClobber] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Append
```powershell
Export-OfficeExcel [-Path] <string> -Append [-InputObject <Object>] [-Sheet <OfficeSheetSelector>] [-TableName <string>] [-Title <string>] [-Address <string>] [-Theme <OfficeThemeInput>] [-TableStyle <ExcelTableStyle>] [-NoTable] [-NoTableStyle] [-NoHeader] [-NoAutoFilter] [-NoAutoFit] [-NoFreeze] [-FreezeFirstColumn] [-ShowFirstColumn] [-ShowLastColumn] [-ShowColumnStripes] [-NoRowStripes] [-ColumnFormat <hashtable>] [-Decimals <int>] [-Culture <cultureinfo>] [-IgnoreMissingColumnFormat] [-ExcludeProperty <string[]>] [-CollectionSeparator <string>] [-FlattenOptions <OfficeObjectFlattenOptions>] [-DocumentProperty <hashtable>] [-DateSystem <ExcelDateSystem>] [-SaveOptions <ExcelSaveOptions>] [-Show] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### ReplaceSheet
```powershell
Export-OfficeExcel [-Path] <string> -ReplaceSheet [-InputObject <Object>] [-Sheet <OfficeSheetSelector>] [-TableName <string>] [-Title <string>] [-Address <string>] [-Theme <OfficeThemeInput>] [-TableStyle <ExcelTableStyle>] [-NoTable] [-NoTableStyle] [-NoHeader] [-NoAutoFilter] [-NoAutoFit] [-NoFreeze] [-FreezeFirstColumn] [-ShowFirstColumn] [-ShowLastColumn] [-ShowColumnStripes] [-NoRowStripes] [-ColumnFormat <hashtable>] [-Decimals <int>] [-Culture <cultureinfo>] [-IgnoreMissingColumnFormat] [-ExcludeProperty <string[]>] [-CollectionSeparator <string>] [-FlattenOptions <OfficeObjectFlattenOptions>] [-DocumentProperty <hashtable>] [-DateSystem <ExcelDateSystem>] [-SaveOptions <ExcelSaveOptions>] [-Show] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Provides a fast PowerShell export path while keeping OfficeIMO as the workbook engine.

## EXAMPLES

### EXAMPLE 1
```powershell
Get-Process | Export-OfficeExcel -Path .\Processes.xlsx -Show
```

Writes the running processes as a styled table with a frozen header and fitted columns, and opens the file.

### EXAMPLE 2
```powershell
$export = @{
    Path         = '.\Reports\Q3-Sales.xlsx'
    Sheet        = 'Orders'
    TableName    = 'Orders'
    Title        = 'Q3 orders'
    ColumnFormat = @{ Margin = 'Percent' }
}
Import-OfficeCsv -Path .\Data\Q3-Orders.csv |
    Select-Object OrderId, Customer, Region, OrderDate, Amount, Margin |
    Export-OfficeExcel @export
```

Exports the Q3 orders as a named table under a title; dates and amounts are formatted from the data, and only the margin column is set explicitly as a percentage.

### EXAMPLE 3
```powershell
Get-Volume |
                Where-Object DriveLetter |
                Select-Object @{ Name = 'Date'; Expression = { Get-Date } }, DriveLetter, Size, SizeRemaining |
                Export-OfficeExcel -Path .\Reports\Disk-Usage.xlsx -Sheet History -TableName DiskHistory -Append
```

Adds today's free-space readings to the DiskHistory table, so the workbook grows into a capacity history one run at a time.

### EXAMPLE 4
```powershell
Import-OfficeCsv -Path .\Exports\open-tickets.csv |
                Export-OfficeExcel -Path .\Reports\Helpdesk.xlsx -Sheet 'Open tickets' -ReplaceSheet
```

Replaces only the Open tickets sheet with this morning's export and leaves the other sheets, charts and notes in the workbook untouched.

## PARAMETERS

### -Address
A1 address of the target cell.

```yaml
Type: String
Parameter Sets: Create, Append, ReplaceSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Append
Append rows to an existing worksheet when the workbook exists.

```yaml
Type: SwitchParameter
Parameter Sets: Append
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -CollectionSeparator
Text inserted between collection values flattened into one cell.

```yaml
Type: String
Parameter Sets: Create, Append, ReplaceSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ColumnFormat
Maps column headers to number-format presets, format codes or per-column option hashtables.

```yaml
Type: Hashtable
Parameter Sets: Create, Append, ReplaceSheet
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
Parameter Sets: Create, Append, ReplaceSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DateSystem
Workbook date system for Excel date serials.

```yaml
Type: ExcelDateSystem
Parameter Sets: Create, Append, ReplaceSheet
Aliases: None
Possible values: NineteenHundred, NineteenFour

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Decimals
Decimal places used for decimal, percent, currency, and scientific presets.

```yaml
Type: Int32
Parameter Sets: Create, Append, ReplaceSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DocumentProperty
Workbook metadata keyed by native property name, such as Title, Author or Company.

```yaml
Type: Hashtable
Parameter Sets: Create, Append, ReplaceSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ExcludeProperty
Exclude specific properties from exported objects.

```yaml
Type: String[]
Parameter Sets: Create, Append, ReplaceSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FlattenOptions
Nested-value projection policy; accepts a native options object or hashtable.

```yaml
Type: OfficeObjectFlattenOptions
Parameter Sets: Create, Append, ReplaceSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FreezeFirstColumn
Freeze the first exported column.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ReplaceSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IgnoreMissingColumnFormat
Continue when a requested export-time column format header is missing.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ReplaceSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -InputObject
Objects, dictionaries or native tabular data supplied directly or through the pipeline.

```yaml
Type: Object
Parameter Sets: Create, Append, ReplaceSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -NoAutoFilter
Omits table filter buttons.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ReplaceSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoAutoFit
Retains native column widths without fitting them to the data.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ReplaceSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoClobber
Do not overwrite an existing workbook.

```yaml
Type: SwitchParameter
Parameter Sets: Create
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoFreeze
Keeps the header row scrollable.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ReplaceSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoHeader
Writes data without a header row.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ReplaceSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoRowStripes
Disables alternating row bands. An explicit false value enables them.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ReplaceSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoTable
Writes plain cells without creating a native Excel table.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ReplaceSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoTableStyle
Creates the table without a named visual style.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ReplaceSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emits the native result after the operation.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ReplaceSheet
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
Type: String
Parameter Sets: Create, Append, ReplaceSheet
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ReplaceSheet
Replaces the selected worksheet while retaining other workbook content.

```yaml
Type: SwitchParameter
Parameter Sets: ReplaceSheet
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SaveOptions
Native save options or a hashtable, applied to a completed content block.

```yaml
Type: ExcelSaveOptions
Parameter Sets: Create, Append, ReplaceSheet
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
Type: OfficeSheetSelector
Parameter Sets: Create, Append, ReplaceSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Show
Open the saved file in its registered application after a local block completes.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ReplaceSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ShowColumnStripes
Enables alternating column bands; explicit false clears them.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ReplaceSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ShowFirstColumn
Emphasizes the first table column; explicit false removes the emphasis.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ReplaceSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ShowLastColumn
Emphasizes the last table column; explicit false removes the emphasis.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ReplaceSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TableName
Optional Excel table name.

```yaml
Type: String
Parameter Sets: Create, Append, ReplaceSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TableStyle
Built-in Excel table style name.

```yaml
Type: ExcelTableStyle
Parameter Sets: Create, Append, ReplaceSheet
Aliases: None
Possible values: TableStyleLight1, TableStyleLight2, TableStyleLight3, TableStyleLight4, TableStyleLight5, TableStyleLight6, TableStyleLight7, TableStyleLight8, TableStyleLight9, TableStyleLight10, TableStyleLight11, TableStyleLight12, TableStyleLight13, TableStyleLight14, TableStyleLight15, TableStyleLight16, TableStyleLight17, TableStyleLight18, TableStyleLight19, TableStyleLight20, TableStyleLight21, TableStyleMedium1, TableStyleMedium2, TableStyleMedium3, TableStyleMedium4, TableStyleMedium5, TableStyleMedium6, TableStyleMedium7, TableStyleMedium8, TableStyleMedium9, TableStyleMedium10, TableStyleMedium11, TableStyleMedium12, TableStyleMedium13, TableStyleMedium14, TableStyleMedium15, TableStyleMedium16, TableStyleMedium17, TableStyleMedium18, TableStyleMedium19, TableStyleMedium20, TableStyleMedium21, TableStyleMedium22, TableStyleMedium23, TableStyleMedium24, TableStyleMedium25, TableStyleMedium26, TableStyleMedium27, TableStyleMedium28, TableStyleDark1, TableStyleDark2, TableStyleDark3, TableStyleDark4, TableStyleDark5, TableStyleDark6, TableStyleDark7, TableStyleDark8, TableStyleDark9, TableStyleDark10, TableStyleDark11, None

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Theme
Built-in Office theme, native theme object or a document whose visual theme is read.

```yaml
Type: OfficeThemeInput
Parameter Sets: Create, Append, ReplaceSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Title
Write a title above the exported table.

```yaml
Type: String
Parameter Sets: Create, Append, ReplaceSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.Object`

## OUTPUTS

- `None`

## RELATED LINKS

- None
