---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficeExcelReportTable
## SYNOPSIS
Adds a titled, styled table to the report.

## SYNTAX
### Default (Default)
```powershell
Add-OfficeExcelReportTable [-InputObject] <Object> [[-Title] <string>] [-Name <string>] [-TableStyle <ExcelTableStyle>] [-Property <string[]>] [-ExcludeProperty <string[]>] [-ColumnFormat <hashtable>] [-CollectionSeparator <string>] [-FlattenOptions <OfficeObjectFlattenOptions>] [-ShowFirstColumn] [-ShowLastColumn] [-NoRowStripes] [-ShowColumnStripes] [-NoAutoFilter] [-NoFreeze] [-NoAutoFormatDynamicCollections] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a titled, styled table to the report.

## EXAMPLES

### EXAMPLE 1
```powershell
New-OfficeExcel -Path .\Reports\Operations.xlsx {
    ExcelReportSheet 'Summary' {
        ExcelReportTable (Get-Service | Where-Object Status -eq 'Stopped') 'Stopped services'
    }
}
```

Adds a styled table with a title to the report.

### EXAMPLE 2
```powershell
$assets = Import-OfficeCsv -Path .\Data\Assets.csv
New-OfficeExcel -Path .\Reports\Asset-Review.xlsx {
    ExcelReportSheet 'Assets' {
        ExcelReportTitle 'Laptops due for replacement'
        ExcelReportTable ($assets | Where-Object AgeYears -ge 4) 'Due this year' -Property AssetTag, User, Model, Purchased, AgeYears -Name DueLaptops -ColumnFormat @{ Purchased = 'Date'; AgeYears = '0.0' }
    }
}
```

Shows only the columns readers need from a wide asset export and leaves the internal identifiers out.

## PARAMETERS

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

### -ExcludeProperty
Exclude specific properties from exported objects.

```yaml
Type: String[]
Parameter Sets: Default
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
Named range name.

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

### -NoAutoFormatDynamicCollections
Disable composer auto-formatting for dynamic collection columns.

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

### -Property
Properties to include, in the requested column order.

```yaml
Type: String[]
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

### -Title
Chart title.

```yaml
Type: String
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: 1
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
