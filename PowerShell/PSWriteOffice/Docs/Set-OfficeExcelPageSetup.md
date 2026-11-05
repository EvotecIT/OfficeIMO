---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Set-OfficeExcelPageSetup
## SYNOPSIS
Sets orientation, margins, scaling, print area and print titles, alone or from a print preset.

## SYNTAX
### Default (Default)
```powershell
Set-OfficeExcelPageSetup [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-Preset <ExcelPrintLayoutPreset>] [-Orientation <OfficePageOrientation>] [-MarginPreset <ExcelMarginPreset>] [-MarginTop <OfficeLength>] [-MarginBottom <OfficeLength>] [-MarginLeft <OfficeLength>] [-MarginRight <OfficeLength>] [-MarginHeader <OfficeLength>] [-MarginFooter <OfficeLength>] [-FitToWidth <UInt32>] [-FitToHeight <UInt32>] [-Scale <UInt32>] [-PageOrder <ExcelPageOrder>] [-PrintArea <string>] [-RepeatRows <string>] [-RepeatColumns <string>] [-NoPrintTitles] [-NoPresetPrintTitles] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Sets orientation, margins, scaling, print area and print titles, alone or from a print preset.

## EXAMPLES

### EXAMPLE 1
```powershell
Open-OfficeExcel -Path .\Reports\Q3-Sales.xlsx {
    ExcelSheet 'Orders' {
        ExcelPageSetup -Orientation Landscape -FitToWidth 1 -FitToHeight 0
    }
}
```

Sets landscape orientation and fits the columns to one page width.

### EXAMPLE 2
```powershell
Open-OfficeExcel -Path .\Reports\Inventory.xlsx {
    ExcelSheet 'Servers' {
        ExcelPageSetup -Preset DataTable -PrintArea A1:H420 -RepeatRows 1:1 -MarginPreset Narrow
    }
}
```

Starts from the data-table print preset, repeats the header row on every page and limits printing to the table area with narrow margins.

## PARAMETERS

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

### -FitToHeight
Number of pages to fit vertically.

```yaml
Type: UInt32
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FitToWidth
Number of pages to fit horizontally.

```yaml
Type: UInt32
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MarginBottom
Bottom page margin in PDF points.

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

### -MarginFooter
Physical distance from the page edge to the footer.

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

### -MarginHeader
Physical distance from the page edge to the header.

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

### -MarginLeft
Left page margin in PDF points.

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

### -MarginPreset
Native print-margin preset applied before explicit margin overrides.

```yaml
Type: ExcelMarginPreset
Parameter Sets: Default
Aliases: None
Possible values: Normal, Narrow, Moderate, Wide

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MarginRight
Right page margin in PDF points.

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

### -MarginTop
Top page margin in PDF points.

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

### -NoPresetPrintTitles
Do not apply print-title rows from the selected preset.

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

### -NoPrintTitles
Clears repeated print rows and columns.

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

### -Orientation
Orientation to apply.

```yaml
Type: OfficePageOrientation
Parameter Sets: Default
Aliases: None
Possible values: Portrait, Landscape

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PageOrder
Multi-page print order.

```yaml
Type: ExcelPageOrder
Parameter Sets: Default
Aliases: None
Possible values: DownThenOver, OverThenDown

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

### -Preset
Margin preset to apply.

```yaml
Type: ExcelPrintLayoutPreset
Parameter Sets: Default
Aliases: None
Possible values: Worksheet, Report, Dashboard, DataTable

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PrintArea
Optional print area in A1 notation.

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

### -RepeatColumns
Inclusive print-title column span, such as A:B.

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

### -RepeatRows
Inclusive print-title row span, such as 1:2.

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

### -Scale
Manual scale percentage (10-400).

```yaml
Type: UInt32
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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `None`

## RELATED LINKS

- None
