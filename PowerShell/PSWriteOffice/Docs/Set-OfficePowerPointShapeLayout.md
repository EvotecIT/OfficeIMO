---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Set-OfficePowerPointShapeLayout
## SYNOPSIS
Aligns, distributes, or arranges PowerPoint shapes using OfficeIMO layout helpers.

## SYNTAX
### Alignment (Default)
```powershell
Set-OfficePowerPointShapeLayout [-InputObject] <Object[]> -Alignment <PowerPointShapeAlignment> [-Slide <PowerPointSlideSelector>] [-ToSlide] [-Margin <OfficeLength>] [-PassThru] [<CommonParameters>]
```

### Distribute
```powershell
Set-OfficePowerPointShapeLayout [-InputObject] <Object[]> -Distribute <PowerPointShapeDistribution> [-Slide <PowerPointSlideSelector>] [-CrossAxisAlignment <PowerPointShapeAlignment>] [-Spacing <OfficeLength>] [-ToSlide] [-Margin <OfficeLength>] [-Center] [-PassThru] [<CommonParameters>]
```

### Grid
```powershell
Set-OfficePowerPointShapeLayout [-InputObject] <Object[]> -Grid [-Slide <PowerPointSlideSelector>] [-Columns <Int32>] [-Rows <Int32>] [-AutoGrid] [-GutterX <OfficeLength>] [-GutterY <OfficeLength>] [-Flow <PowerPointShapeGridFlow>] [-NoResize] [-ToSlide] [-Margin <OfficeLength>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Aligns, distributes, or arranges PowerPoint shapes using OfficeIMO layout helpers.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Get-OfficePowerPointShape -Slide $slide -Name 'Kpi.*' |
                Set-OfficePowerPointShapeLayout -Alignment Top
```

Uses OfficeIMO.PowerPoint to align all matching shapes to the top edge of their selection bounds.

## PARAMETERS

### -Alignment
Alignment operation.

```yaml
Type: PowerPointShapeAlignment
Parameter Sets: Alignment
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

### -CrossAxisAlignment
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

### -GutterX
Horizontal grid gutter; bare numbers mean points.

```yaml
Type: OfficeLength
Parameter Sets: Grid
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -GutterY
Vertical grid gutter; bare numbers mean points.

```yaml
Type: OfficeLength
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
PowerPoint shapes or shape info records from Get-OfficePowerPointShape.

```yaml
Type: Object[]
Parameter Sets: Alignment, Distribute, Grid
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Margin
Use slide content bounds with the supplied margin; bare numbers mean points.

```yaml
Type: OfficeLength
Parameter Sets: Alignment, Distribute, Grid
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
Parameter Sets: Alignment, Distribute, Grid
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
Native slide or one-based slide number; omission uses the supplied shapes' owning slide.

```yaml
Type: PowerPointSlideSelector
Parameter Sets: Alignment, Distribute, Grid
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Spacing
Fixed spacing between distributed shapes; bare numbers mean points.

```yaml
Type: OfficeLength
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
Parameter Sets: Alignment, Distribute, Grid
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

- `System.Object[]`

## OUTPUTS

- `OfficeIMO.PowerPoint.PowerPointShape`

## RELATED LINKS

- None
