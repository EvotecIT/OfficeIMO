---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficePdfExcelOptions
## SYNOPSIS
Creates discoverable PDF-table-to-Excel reconstruction settings.

## SYNTAX
### Default (Default)
```powershell
New-OfficePdfExcelOptions [-MaxRows <Int32>] [-SheetNamePrefix <string>] [-TableNamePrefix <string>] [-TableStyle <ExcelTableStyle>] [-NoAutoFilter] [-NoAutoFit] [-NoNumericConversion] [-NoBooleanConversion] [-NoDateTimeConversion] [-NoPercentageConversion] [-Culture <cultureinfo>] [-NoPageContinuationMerge] [-KeepRepeatedHeaderRows] [-MaximumContinuationSegments <Int32>] [-ContinuationGeometryTolerance <OfficeLength>] [-EmptyWorkbookSheetName <string>] [<CommonParameters>]
```

## DESCRIPTION
Creates discoverable PDF-table-to-Excel reconstruction settings.

## EXAMPLES

### EXAMPLE 1
```powershell
$options = New-OfficePdfExcelOptions -Culture en-GB
Convert-OfficeDocument -Path ./Tables.pdf -To Excel -OutputPath ./Tables.xlsx -Options $options
```


### EXAMPLE 2
```powershell
$options = New-OfficePdfExcelOptions -NoNumericConversion -NoDateTimeConversion -NoAutoFit
Convert-OfficeDocument -Path ./Tables.pdf -To Excel -OutputPath ./Text.xlsx -Options $options
```


## PARAMETERS

### -ContinuationGeometryTolerance
Physical geometry tolerance for page continuations; bare numbers mean points.

```yaml
Type: OfficeLength
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
Culture used for numeric parsing; defaults to the current culture.

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

### -EmptyWorkbookSheetName
Worksheet name used when no tables are detected.

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

### -KeepRepeatedHeaderRows
Retain repeated body header rows in merged table segments.

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

### -MaximumContinuationSegments
Maximum table segments merged into one table.

```yaml
Type: Int32
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxRows
Maximum body rows imported per detected table; zero means unlimited.

```yaml
Type: Int32
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
Exclude table-scoped AutoFilters.

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
Keep native widths without fitting worksheet columns.

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

### -NoBooleanConversion
Keep boolean-looking values as text.

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

### -NoDateTimeConversion
Keep date-looking values as text.

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

### -NoNumericConversion
Keep numeric-looking values as text.

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

### -NoPageContinuationMerge
Keep table segments on separate worksheets instead of merging across pages.

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

### -NoPercentageConversion
Keep percentage values as text.

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

### -SheetNamePrefix
Prefix for generated worksheet names.

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

### -TableNamePrefix
Prefix for generated Excel table names.

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

### -TableStyle
Excel table style.

```yaml
Type: ExcelTableStyle
Parameter Sets: Default
Aliases: None
Possible values: TableStyleLight1, TableStyleLight2, TableStyleLight3, TableStyleLight4, TableStyleLight5, TableStyleLight6, TableStyleLight7, TableStyleLight8, TableStyleLight9, TableStyleLight10, TableStyleLight11, TableStyleLight12, TableStyleLight13, TableStyleLight14, TableStyleLight15, TableStyleLight16, TableStyleLight17, TableStyleLight18, TableStyleLight19, TableStyleLight20, TableStyleLight21, TableStyleMedium1, TableStyleMedium2, TableStyleMedium3, TableStyleMedium4, TableStyleMedium5, TableStyleMedium6, TableStyleMedium7, TableStyleMedium8, TableStyleMedium9, TableStyleMedium10, TableStyleMedium11, TableStyleMedium12, TableStyleMedium13, TableStyleMedium14, TableStyleMedium15, TableStyleMedium16, TableStyleMedium17, TableStyleMedium18, TableStyleMedium19, TableStyleMedium20, TableStyleMedium21, TableStyleMedium22, TableStyleMedium23, TableStyleMedium24, TableStyleMedium25, TableStyleMedium26, TableStyleMedium27, TableStyleMedium28, TableStyleDark1, TableStyleDark2, TableStyleDark3, TableStyleDark4, TableStyleDark5, TableStyleDark6, TableStyleDark7, TableStyleDark8, TableStyleDark9, TableStyleDark10, TableStyleDark11

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `OfficeIMO.Pdf.Excel.ExcelPdfImportOptions`

## RELATED LINKS

- None
