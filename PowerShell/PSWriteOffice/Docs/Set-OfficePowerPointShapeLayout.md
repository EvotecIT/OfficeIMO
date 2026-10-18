---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficePowerPointShapeLayout
## SYNOPSIS
Aligns, distributes, or arranges PowerPoint shapes using OfficeIMO layout helpers.

## SYNTAX
### Align (Default)
```powershell
Set-OfficePowerPointShapeLayout [-InputObject] <Object> -Align <PowerPointShapeAlignment> [-Slide <PowerPointSlide>] [-ToSlide] [-MarginPoints <Double>] [-PassThru] [<CommonParameters>]
```

### Distribute
```powershell
Set-OfficePowerPointShapeLayout [-InputObject] <Object> -Distribute <PowerPointShapeDistribution> [-Slide <PowerPointSlide>] [-CrossAxisAlign <PowerPointShapeAlignment>] [-SpacingPoints <Double>] [-ToSlide] [-MarginPoints <Double>] [-Center] [-PassThru] [<CommonParameters>]
```

### Grid
```powershell
Set-OfficePowerPointShapeLayout [-InputObject] <Object> -Grid [-Slide <PowerPointSlide>] [-Columns <Int32>] [-Rows <Int32>] [-AutoGrid] [-GutterXPoints <double>] [-GutterYPoints <double>] [-Flow <PowerPointShapeGridFlow>] [-NoResize] [-ToSlide] [-MarginPoints <Double>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Aligns, distributes, or arranges PowerPoint shapes using OfficeIMO layout helpers.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Find-OfficePowerPointShape -Slide $slide -Name 'Kpi.*' |
                Set-OfficePowerPointShapeLayout -Align Top
```

Uses OfficeIMO.PowerPoint to align all matching shapes to the top edge of their selection bounds.

## PARAMETERS

### -Align
Alignment operation.

```yaml
Type: PowerPointShapeAlignment
Parameter Sets: Align
Aliases: None
Possible values: Left, Center, Right, Top, Middle, Bottom

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AutoGrid
Let OfficeIMO choose the grid dimensions.

```yaml
Type: SwitchParameter
Parameter Sets: Grid
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Center
Center a fixed-spacing distribution within its bounds.

```yaml
Type: SwitchParameter
Parameter Sets: Distribute
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Columns
Grid column count. Omit with AutoGrid.

```yaml
Type: Int32
Parameter Sets: Grid
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -CrossAxisAlign
Optional cross-axis alignment for even distribution.

```yaml
Type: PowerPointShapeAlignment
Parameter Sets: Distribute
Aliases: None
Possible values: Left, Center, Right, Top, Middle, Bottom

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Distribute
Distribution operation.

```yaml
Type: PowerPointShapeDistribution
Parameter Sets: Distribute
Aliases: None
Possible values: Horizontal, Vertical

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Flow
Fill the grid column-by-column instead of row-by-row.

```yaml
Type: PowerPointShapeGridFlow
Parameter Sets: Grid
Aliases: None
Possible values: RowMajor, ColumnMajor

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Grid
Arrange shapes in a grid.

```yaml
Type: SwitchParameter
Parameter Sets: Grid
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -GutterXPoints
Horizontal grid gutter in points.

```yaml
Type: Double
Parameter Sets: Grid
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -GutterYPoints
Vertical grid gutter in points.

```yaml
Type: Double
Parameter Sets: Grid
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -InputObject
PowerPoint shapes or shape info records from Get-OfficePowerPointShape or Find-OfficePowerPointShape.

```yaml
Type: Object
Parameter Sets: Align, Distribute, Grid
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -MarginPoints
Use slide content bounds with the supplied margin in points.

```yaml
Type: Double
Parameter Sets: Align, Distribute, Grid
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoResize
Keep each shape's current size when arranging in a grid.

```yaml
Type: SwitchParameter
Parameter Sets: Grid
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the arranged shapes.

```yaml
Type: SwitchParameter
Parameter Sets: Align, Distribute, Grid
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Rows
Grid row count. Omit with AutoGrid.

```yaml
Type: Int32
Parameter Sets: Grid
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Slide
Slide that owns raw PowerPointShape inputs. Shape info records carry their own slide.

```yaml
Type: PowerPointSlide
Parameter Sets: Align, Distribute, Grid
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SpacingPoints
Fixed spacing between distributed shapes in points.

```yaml
Type: Double
Parameter Sets: Distribute
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ToSlide
Use the full slide bounds instead of the current selection bounds.

```yaml
Type: SwitchParameter
Parameter Sets: Align, Distribute, Grid
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

- `OfficeIMO.PowerPoint.PowerPointShape`

## RELATED LINKS

- None
