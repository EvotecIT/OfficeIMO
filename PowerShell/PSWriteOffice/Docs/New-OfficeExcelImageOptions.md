---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficeExcelImageOptions
## SYNOPSIS
Creates Excel rendering settings for Export-OfficeDocumentImage and ConvertTo-OfficeDocumentImage.

## SYNTAX
### Default (Default)
```powershell
New-OfficeExcelImageOptions [-ShowGridlines] [-IncludeHidden] [-IncludeImages] [-IncludeCharts] [-IncludeDrawingObjects] [-IncludeConditionalFormatting] [-ShowHyperlinkHints] [-ShowCommentBodies] [-MaximumRenderedCells <Int32>] [-MaximumOutputWidth <Int32>] [-MaximumOutputHeight <Int32>] [-MaximumRasterPixelCount <Int64>] [-RasterOverflowBehavior <OfficeRasterOverflowBehavior>] [-MaximumOutputCount <Int32>] [-MaximumTotalRasterPixelCount <Int64>] [-MaximumTotalEncodedBytes <Int64>] [-RenderTimeoutSeconds <Double>] [-MaximumDegreeOfParallelism <Int32>] [-TextShapingLanguage <string>] [-IncludeHiddenSheets] [-UseWorksheetPrintAreas] [-SplitWorksheetsByManualPageBreaks] [<CommonParameters>]
```

## DESCRIPTION
Creates Excel rendering settings for Export-OfficeDocumentImage and ConvertTo-OfficeDocumentImage.

## EXAMPLES

### EXAMPLE 1
```powershell
$image = New-OfficeExcelImageOptions -ShowGridlines
Export-OfficeDocumentImage -Path .\Reports\Q3-Sales.xlsx -Sheet Summary -OutputPath .\Images\Q3-Summary.png -Options $image
```

Renders the Summary sheet to PNG with gridlines shown.

### EXAMPLE 2
```powershell
$image = New-OfficeExcelImageOptions -IncludeCharts -IncludeImages -IncludeConditionalFormatting
Export-OfficeDocumentImage -Path .\Reports\Operations.xlsx -Sheet Dashboard -OutputDirectory .\Site\Images -Scale 2 -Options $image
```

Renders the dashboard at double scale with charts and images, so the PNG stays sharp when embedded in a web page.

## PARAMETERS

### -IncludeCharts
Includes charts in the rendered output; explicit false omits them.

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

### -IncludeConditionalFormatting
Evaluates supported conditional formatting during rendering; explicit false disables it.

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

### -IncludeDrawingObjects
Includes worksheet drawing objects in the rendered output.

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

### -IncludeHidden
Renders hidden rows and columns; explicit false omits them.

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

### -IncludeHiddenSheets
Includes hidden worksheets when no explicit sheet selection is supplied.

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

### -IncludeImages
Includes embedded worksheet images in the rendered output.

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

### -MaximumDegreeOfParallelism
Maximum concurrent render workers; omission retains the native limit.

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

### -MaximumOutputCount
Maximum number of images produced by one batch operation.

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

### -MaximumOutputHeight
Maximum output height in pixels.

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

### -MaximumOutputWidth
Maximum output width in pixels.

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

### -MaximumRasterPixelCount
Maximum pixel count allocated for one raster image.

```yaml
Type: Int64
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumRenderedCells
Maximum worksheet cells materialized for one visual snapshot.

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

### -MaximumTotalEncodedBytes
Maximum aggregate encoded image bytes in one batch.

```yaml
Type: Int64
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumTotalRasterPixelCount
Maximum aggregate raster pixels in one batch.

```yaml
Type: Int64
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RasterOverflowBehavior
Native policy when the requested raster exceeds its pixel budget.

```yaml
Type: OfficeRasterOverflowBehavior
Parameter Sets: Default
Aliases: None
Possible values: ReduceScale, Throw

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RenderTimeoutSeconds
Render timeout in seconds; omission retains the native timeout.

```yaml
Type: Double
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ShowCommentBodies
Renders comment text rather than only comment indicators.

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

### -ShowGridlines
Renders worksheet gridlines; explicit false hides them.

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

### -ShowHyperlinkHints
Renders hyperlink hints in worksheet snapshots.

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

### -SplitWorksheetsByManualPageBreaks
Creates separate images at worksheet manual page breaks.

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

### -TextShapingLanguage
Language tag used for native text shaping.

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

### -UseWorksheetPrintAreas
Honor worksheet print areas.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `None`

## RELATED LINKS

- None
