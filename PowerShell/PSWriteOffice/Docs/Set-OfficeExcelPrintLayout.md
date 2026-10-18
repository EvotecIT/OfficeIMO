---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeExcelPrintLayout
## SYNOPSIS
Applies a reusable worksheet print layout preset.

## SYNTAX
### Context (Default)
```powershell
Set-OfficeExcelPrintLayout [-Sheet <string>] [-SheetIndex <Int32>] [-Preset <ExcelPrintLayoutPreset>] [-PrintArea <string>] [-Orientation <OfficePageOrientation>] [-Margins <ExcelMarginPreset>] [-FitToWidth <UInt32>] [-FitToHeight <UInt32>] [-Scale <UInt32>] [-PageOrder <ExcelPageOrder>] [-RepeatFirstRow <Int32>] [-RepeatLastRow <Int32>] [-RepeatFirstColumn <Int32>] [-RepeatLastColumn <Int32>] [-NoPresetPrintTitles] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Path
```powershell
Set-OfficeExcelPrintLayout [-Path] <string> [-Sheet <string>] [-SheetIndex <Int32>] [-Preset <ExcelPrintLayoutPreset>] [-PrintArea <string>] [-Orientation <OfficePageOrientation>] [-Margins <ExcelMarginPreset>] [-FitToWidth <UInt32>] [-FitToHeight <UInt32>] [-Scale <UInt32>] [-PageOrder <ExcelPageOrder>] [-RepeatFirstRow <Int32>] [-RepeatLastRow <Int32>] [-RepeatFirstColumn <Int32>] [-RepeatLastColumn <Int32>] [-NoPresetPrintTitles] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
Set-OfficeExcelPrintLayout -Document <ExcelDocument> [-Sheet <string>] [-SheetIndex <Int32>] [-Preset <ExcelPrintLayoutPreset>] [-PrintArea <string>] [-Orientation <OfficePageOrientation>] [-Margins <ExcelMarginPreset>] [-FitToWidth <UInt32>] [-FitToHeight <UInt32>] [-Scale <UInt32>] [-PageOrder <ExcelPageOrder>] [-RepeatFirstRow <Int32>] [-RepeatLastRow <Int32>] [-RepeatFirstColumn <Int32>] [-RepeatLastColumn <Int32>] [-NoPresetPrintTitles] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Applies a reusable worksheet print layout preset.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> ExcelSheet 'Report' { Set-OfficeExcelPrintLayout -Preset Report -PrintArea A1:H40 }
```

Applies landscape orientation, narrow margins, one-page-wide scaling, and repeated header row.

## PARAMETERS

### -Document
Workbook to update outside the DSL context.

```yaml
Type: ExcelDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -FitToHeight
Optional pages-tall fit override. Use 0 for unlimited height.

```yaml
Type: UInt32
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FitToWidth
Optional pages-wide fit override.

```yaml
Type: UInt32
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Margins
Optional margin preset override.

```yaml
Type: ExcelMarginPreset
Parameter Sets: Context, Path, Document
Aliases: None
Possible values: Normal, Narrow, Moderate, Wide

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
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Orientation
Optional orientation override.

```yaml
Type: OfficePageOrientation
Parameter Sets: Context, Path, Document
Aliases: None
Possible values: Portrait, Landscape

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PageOrder
Optional multi-page print order override.

```yaml
Type: ExcelPageOrder
Parameter Sets: Context, Path, Document
Aliases: None
Possible values: DownThenOver, OverThenDown

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the worksheet after applying the layout.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Workbook path to update.

```yaml
Type: String
Parameter Sets: Path
Aliases: InputPath, FilePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Preset
Print layout preset.

```yaml
Type: ExcelPrintLayoutPreset
Parameter Sets: Context, Path, Document
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
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RepeatFirstColumn
Optional first 1-based repeated print-title column.

```yaml
Type: Int32
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RepeatFirstRow
Optional first 1-based repeated print-title row.

```yaml
Type: Int32
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RepeatLastColumn
Optional last 1-based repeated print-title column.

```yaml
Type: Int32
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RepeatLastRow
Optional last 1-based repeated print-title row.

```yaml
Type: Int32
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Scale
Optional manual scale percentage override.

```yaml
Type: UInt32
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Sheet
Worksheet name. Defaults to the current sheet inside an ExcelSheet block.

```yaml
Type: String
Parameter Sets: Context, Path, Document
Aliases: WorksheetName
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SheetIndex
Worksheet index when using a workbook object or path.

```yaml
Type: Int32
Parameter Sets: Context, Path, Document
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

- `OfficeIMO.Excel.ExcelSheet`
- `System.Management.Automation.PSObject`

## RELATED LINKS

- None
