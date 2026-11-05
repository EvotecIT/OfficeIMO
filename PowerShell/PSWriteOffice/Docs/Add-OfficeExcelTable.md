---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficeExcelTable
## SYNOPSIS
Writes tabular data to the current worksheet and formats it as an Excel table.

## SYNTAX
### Default (Default)
```powershell
Add-OfficeExcelTable [-InputObject] <Object> [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-Name <string>] [-Address <string>] [-TableStyle <ExcelTableStyle>] [-NoTable] [-NoTableStyle] [-NoHeader] [-NoAutoFilter] [-NoAutoFit] [-NoFreeze] [-ShowFirstColumn] [-ShowLastColumn] [-ShowColumnStripes] [-NoRowStripes] [-ColumnFormat <hashtable>] [-Culture <cultureinfo>] [-View <OfficeTableView>] [-CollectionSeparator <string>] [-FlattenOptions <OfficeObjectFlattenOptions>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Accepts objects, dictionaries, DataTable/DataView/IDataReader inputs, or DataRow sequences and writes them into an Excel table with optional styling.

## EXAMPLES

### EXAMPLE 1
```powershell
New-OfficeExcel -Path .\Services.xlsx {
    ExcelSheet 'Services' {
        ExcelTable (Get-Service)
    }
}
```

Writes the services as a themed table with a frozen header, filters and fitted columns.

### EXAMPLE 2
```powershell
$targets = Import-OfficeCsv -Path .\Data\Q3-Targets.csv
$actuals = Import-OfficeCsv -Path .\Data\Q3-Actuals.csv
New-OfficeExcel -Path .\Reports\Q3-Performance.xlsx {
    ExcelSheet 'Performance' {
        ExcelTable $targets -Name Targets -Address A1
        ExcelTable $actuals -Name Actuals -Address G1 -ColumnFormat @{ Attainment = 'Percent' }
    }
}
```

Writes this quarter's targets and actuals side by side as named tables, so charts, formulas and later imports can refer to them by name.

## PARAMETERS

### -Address
A1 address of the target cell.

```yaml
Type: String
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -CollectionSeparator
Text inserted between collection values flattened into one cell.

```yaml
Type: String
Parameter Sets: Default
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
Parameter Sets: Default
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
Parameter Sets: Default
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
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -FlattenOptions
Nested-value projection policy; accepts a native options object or hashtable.

```yaml
Type: OfficeObjectFlattenOptions
Parameter Sets: Default
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
Parameter Sets: Default
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Name
Name to assign to the table.

```yaml
Type: String
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoAutoFilter
Omits table filter buttons.

```yaml
Type: SwitchParameter
Parameter Sets: Default
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
Parameter Sets: Default
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
Parameter Sets: Default
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
Parameter Sets: Default
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
Parameter Sets: Default
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
Parameter Sets: Default
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
Parameter Sets: Default
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
Parameter Sets: Default
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
Parameter Sets: Default
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
Parameter Sets: Default
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
Parameter Sets: Default
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
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TableStyle
Built-in table style to apply.

```yaml
Type: ExcelTableStyle
Parameter Sets: Default
Aliases: None
Possible values: TableStyleLight1, TableStyleLight2, TableStyleLight3, TableStyleLight4, TableStyleLight5, TableStyleLight6, TableStyleLight7, TableStyleLight8, TableStyleLight9, TableStyleLight10, TableStyleLight11, TableStyleLight12, TableStyleLight13, TableStyleLight14, TableStyleLight15, TableStyleLight16, TableStyleLight17, TableStyleLight18, TableStyleLight19, TableStyleLight20, TableStyleLight21, TableStyleMedium1, TableStyleMedium2, TableStyleMedium3, TableStyleMedium4, TableStyleMedium5, TableStyleMedium6, TableStyleMedium7, TableStyleMedium8, TableStyleMedium9, TableStyleMedium10, TableStyleMedium11, TableStyleMedium12, TableStyleMedium13, TableStyleMedium14, TableStyleMedium15, TableStyleMedium16, TableStyleMedium17, TableStyleMedium18, TableStyleMedium19, TableStyleMedium20, TableStyleMedium21, TableStyleMedium22, TableStyleMedium23, TableStyleMedium24, TableStyleMedium25, TableStyleMedium26, TableStyleMedium27, TableStyleMedium28, TableStyleDark1, TableStyleDark2, TableStyleDark3, TableStyleDark4, TableStyleDark5, TableStyleDark6, TableStyleDark7, TableStyleDark8, TableStyleDark9, TableStyleDark10, TableStyleDark11, None

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -View
Projection to apply before writing the table.

```yaml
Type: OfficeTableView
Parameter Sets: Default
Aliases: None
Possible values: Normal, Transpose

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Excel.ExcelDocument`
- `System.Object`

## OUTPUTS

- `None`

## RELATED LINKS

- None
