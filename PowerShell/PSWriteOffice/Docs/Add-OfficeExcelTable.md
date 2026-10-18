---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeExcelTable
## SYNOPSIS
Writes tabular data to the current worksheet and formats it as an Excel table.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficeExcelTable [-InputObject] <Object> [-Worksheet <ExcelSheet>] [-Document <ExcelDocument>] [-Sheet <string>] [-SheetIndex <Int32>] [-StartRow <int>] [-StartColumn <int>] [-NoHeader] [-View <OfficeTableView>] [-CollectionSeparator <string>] [-DictionaryEntrySeparator <string>] [-DictionaryKeyValueSeparator <string>] [-MaxCollectionItems <int>] [-MaxNestingDepth <int>] [-TableName <string>] [-TableStyle <ExcelTableStyle>] [-ShowFirstColumn] [-ShowLastColumn] [-NoRowStripes] [-ShowColumnStripes] [-NoAutoFilter] [-AutoFit] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Accepts objects, dictionaries, DataTable/DataView/IDataReader inputs, or DataRow sequences and writes them into an Excel table with optional styling.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $data = @([pscustomobject]@{ Region='NA'; Revenue=100 }, [pscustomobject]@{ Region='EMEA'; Revenue=150 })
ExcelSheet 'Data' { Add-OfficeExcelTable -InputObject $data -TableName 'Sales' }
```

Writes two rows and formats them as a styled Excel table.

### EXAMPLE 2
```powershell
PS> Add-OfficeExcelTable -Worksheet $sheet -InputObject $rows -TableName 'Sales' -AutoFit
```

Writes the rows into a live workbook without requiring an active DSL scope.

## PARAMETERS

### -AutoFit
Auto-fit the table columns after insertion.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -CollectionSeparator
Text used between items when a cell contains a collection.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DictionaryEntrySeparator
Text used between entries when a cell contains a dictionary.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DictionaryKeyValueSeparator
Text used between a dictionary key and value.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Workbook that will receive the table outside a DSL context.

```yaml
Type: ExcelDocument
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -InputObject
Source objects, dictionaries, DataTable/DataView/IDataReader inputs, or DataRow sequences to convert into table rows.

```yaml
Type: Object
Parameter Sets: __AllParameterSets
Aliases: Data, DataTable
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -MaxCollectionItems
Maximum number of items allowed in one nested collection or dictionary cell. Defaults to 1,048,575; increase explicitly for trusted larger values.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxNestingDepth
Maximum nesting depth allowed while normalizing one cell value. Defaults to 64; increase explicitly for trusted deeper values.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoAutoFilter
Disable AutoFilter dropdowns.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoHeader
Skip writing headers.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoRowStripes
Disable alternating row stripes for the created table.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Return the created range string.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Sheet
Worksheet name when using Document.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SheetIndex
Worksheet index (0-based) when using Document.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ShowColumnStripes
Enable alternating column stripes for the created table.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ShowFirstColumn
Emphasize the first table column when the selected style supports it.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ShowLastColumn
Emphasize the last table column when the selected style supports it.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -StartColumn
Starting column for the data (1-based).

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -StartRow
Starting row for the data (1-based).

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TableName
Name to assign to the table.

```yaml
Type: String
Parameter Sets: __AllParameterSets
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
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: TableStyleLight1, TableStyleLight2, TableStyleLight3, TableStyleLight4, TableStyleLight5, TableStyleLight6, TableStyleLight7, TableStyleLight8, TableStyleLight9, TableStyleLight10, TableStyleLight11, TableStyleLight12, TableStyleLight13, TableStyleLight14, TableStyleLight15, TableStyleLight16, TableStyleLight17, TableStyleLight18, TableStyleLight19, TableStyleLight20, TableStyleLight21, TableStyleMedium1, TableStyleMedium2, TableStyleMedium3, TableStyleMedium4, TableStyleMedium5, TableStyleMedium6, TableStyleMedium7, TableStyleMedium8, TableStyleMedium9, TableStyleMedium10, TableStyleMedium11, TableStyleMedium12, TableStyleMedium13, TableStyleMedium14, TableStyleMedium15, TableStyleMedium16, TableStyleMedium17, TableStyleMedium18, TableStyleMedium19, TableStyleMedium20, TableStyleMedium21, TableStyleMedium22, TableStyleMedium23, TableStyleMedium24, TableStyleMedium25, TableStyleMedium26, TableStyleMedium27, TableStyleMedium28, TableStyleDark1, TableStyleDark2, TableStyleDark3, TableStyleDark4, TableStyleDark5, TableStyleDark6, TableStyleDark7, TableStyleDark8, TableStyleDark9, TableStyleDark10, TableStyleDark11

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
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Normal, Transpose

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Worksheet
Worksheet that will receive the table outside a DSL context.

```yaml
Type: ExcelSheet
Parameter Sets: __AllParameterSets
Aliases: SheetObject
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
