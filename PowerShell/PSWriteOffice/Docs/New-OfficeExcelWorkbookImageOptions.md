---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# New-OfficeExcelWorkbookImageOptions
## SYNOPSIS
Creates discoverable sheet selection and rendering settings for Export-OfficeExcelImage.

## SYNTAX
### __AllParameterSets
```powershell
New-OfficeExcelWorkbookImageOptions [-SheetName <string[]>] [-IncludeHiddenSheets] [-UseWorksheetPrintAreas] [-SplitWorksheetsByManualPageBreaks] [-ShowGridlines] [-IncludeHidden] [-IncludeImages] [-IncludeCharts] [-IncludeDrawingObjects] [-IncludeConditionalFormatting] [-MaximumRenderedCells <Int32>] [-Scale <Double>] [-MaximumOutputWidth <Int32>] [-MaximumOutputHeight <Int32>] [-BackgroundColor <string>] [-TargetDpi <Double>] [-MaximumRasterPixels <Int64>] [-RasterOverflowBehavior <OfficeRasterOverflowBehavior>] [-MaximumOutputCount <Int32>] [-MaximumTotalRasterPixels <Int64>] [-MaximumTotalEncodedBytes <Int64>] [-RenderTimeoutSeconds <Double>] [-MaximumDegreeOfParallelism <Int32>] [-TextShapingLanguage <string>] [<CommonParameters>]
```

## DESCRIPTION
Creates discoverable sheet selection and rendering settings for Export-OfficeExcelImage.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $options = New-OfficeExcelWorkbookImageOptions -SheetName Summary,Data -IncludeCharts -IncludeConditionalFormatting
Export-OfficeExcelImage -Path .\Workbook.xlsx -OutputPath .\Sheets -Options $options
```


## PARAMETERS

### -BackgroundColor
Background color name or hex value.

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

### -IncludeCharts
Include worksheet charts.

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

### -IncludeConditionalFormatting
Include conditional formatting.

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

### -IncludeDrawingObjects
Include drawing objects.

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

### -IncludeHidden
Include hidden rows and columns.

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

### -IncludeHiddenSheets
Include hidden worksheets.

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

### -IncludeImages
Include worksheet images.

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

### -MaximumDegreeOfParallelism
Maximum independent renders processed concurrently.

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

### -MaximumOutputCount
Maximum images accepted from one batch export.

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

### -MaximumOutputHeight
Maximum output height in pixels.

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

### -MaximumOutputWidth
Maximum output width in pixels.

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

### -MaximumRasterPixels
Maximum pixels allocated for one raster image.

```yaml
Type: Int64
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumRenderedCells
Maximum cells rendered per worksheet.

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

### -MaximumTotalEncodedBytes
Maximum aggregate encoded bytes accepted from one batch.

```yaml
Type: Int64
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumTotalRasterPixels
Maximum aggregate raster pixels accepted from one batch.

```yaml
Type: Int64
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RasterOverflowBehavior
Reduce or reject oversized raster output.

```yaml
Type: OfficeRasterOverflowBehavior
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: ReduceScale, Throw

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RenderTimeoutSeconds
Maximum seconds allowed for one render.

```yaml
Type: Double
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Scale
Output scale multiplier.

```yaml
Type: Double
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SheetName
Worksheet names to export.

```yaml
Type: String[]
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ShowGridlines
Show worksheet gridlines.

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

### -SplitWorksheetsByManualPageBreaks
Split worksheets at manual page breaks.

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

### -TargetDpi
Target output density in dots per inch.

```yaml
Type: Double
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TextShapingLanguage
BCP 47 language hint for text shaping.

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

### -UseWorksheetPrintAreas
Use worksheet print areas.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `OfficeIMO.Excel.ExcelWorkbookImageExportOptions`

## RELATED LINKS

- None
