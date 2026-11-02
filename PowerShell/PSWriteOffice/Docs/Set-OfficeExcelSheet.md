---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Set-OfficeExcelSheet
## SYNOPSIS
Sets worksheet visibility, tab color and the active worksheet.

## SYNTAX
### Default (Default)
```powershell
Set-OfficeExcelSheet [-Sheet <OfficeSheetSelector>] [-Visibility <ExcelSheetVisibility>] [-TabColor <OfficeColor>] [-Active] [-PassThru] [-Document <ExcelDocument>] [<CommonParameters>]
```

### NoTabColor
```powershell
Set-OfficeExcelSheet -NoTabColor [-Sheet <OfficeSheetSelector>] [-Visibility <ExcelSheetVisibility>] [-Active] [-PassThru] [-Document <ExcelDocument>] [<CommonParameters>]
```

## DESCRIPTION
Omitted properties keep their values. Sheet accepts a name, one-based position or native worksheet.
Inside an ExcelSheet block it defaults to the current worksheet; an explicit workbook defaults to its first sheet.

## EXAMPLES

### EXAMPLE 1
```powershell
Open-OfficeExcel -Path .\Pricing.xlsx {
    Set-OfficeExcelSheet -Sheet Lookups -Visibility VeryHidden -TabColor Accent2
    Set-OfficeExcelSheet -Sheet Summary -Active
}
```


### EXAMPLE 2
```powershell
$book | Set-OfficeExcelSheet -Sheet 2 -NoTabColor -PassThru
```


## PARAMETERS

### -Active
Gets or sets whether this visible worksheet opens as the workbook's active worksheet.

```yaml
Type: SwitchParameter
Parameter Sets: Default, NoTabColor
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Gets or sets an explicit live document; omission uses the enclosing content block.

```yaml
Type: ExcelDocument
Parameter Sets: Default, NoTabColor
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -NoTabColor
Gets or sets whether to remove the tab color; exclusive with TabColor.

```yaml
Type: SwitchParameter
Parameter Sets: NoTabColor
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Gets or sets whether to emit the edited native worksheet.

```yaml
Type: SwitchParameter
Parameter Sets: Default, NoTabColor
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Sheet
Gets or sets a worksheet name, one-based position or native worksheet belonging to the workbook.

```yaml
Type: OfficeSheetSelector
Parameter Sets: Default, NoTabColor
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TabColor
Gets or sets the literal or theme-based tab color.

```yaml
Type: OfficeColor
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Visibility
Gets or sets Visible, Hidden or VeryHidden; omission preserves visibility.

```yaml
Type: ExcelSheetVisibility
Parameter Sets: Default, NoTabColor
Aliases: None
Possible values: Visible, Hidden, VeryHidden

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

## OUTPUTS

- `OfficeIMO.Excel.ExcelSheet`

## RELATED LINKS

- None
