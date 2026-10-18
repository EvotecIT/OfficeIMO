---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeVisioShapeLayout
## SYNOPSIS
Applies OfficeIMO Visio selection layout and layer operations to shapes.

## SYNTAX
### __AllParameterSets
```powershell
Set-OfficeVisioShapeLayout [[-InputObject] <Object>] [-Page <VisioPage>] [-ShapeId <string[]>] [-Layer <string>] [-AlignHorizontal <VisioHorizontalAlignment>] [-AlignVertical <VisioVerticalAlignment>] [-Distribute <VisioDistributionAxis>] [-Grid] [-HorizontalStack] [-VerticalStack] [-Columns <int>] [-HorizontalSpacing <double>] [-VerticalSpacing <double>] [-PreserveFirstShapeCenter] [-NoRouteInternalConnectors] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Applies OfficeIMO Visio selection layout and layer operations to shapes.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> OfficeVisio -Path .\Layout.vsdx {
    VisioRectangle -Key one -Text 'One'
    VisioRectangle -Key two -Text 'Two'
    VisioLayout -ShapeId one,two -HorizontalStack -HorizontalSpacing 0.4
}
```

Resolves keyed shapes and applies a reusable OfficeIMO Visio layout operation.

## PARAMETERS

### -AlignHorizontal
Horizontal alignment inside the selected shapes' bounds.

```yaml
Type: VisioHorizontalAlignment
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Left, Center, Right

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AlignVertical
Vertical alignment inside the selected shapes' bounds.

```yaml
Type: VisioVerticalAlignment
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Bottom, Middle, Top

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Columns
Grid column count. Zero lets OfficeIMO choose a near-square grid.

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

### -Distribute
Distribute selected shapes along an axis.

```yaml
Type: VisioDistributionAxis
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Horizontal, Vertical

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Grid
Lay out selected shapes as a grid.

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

### -HorizontalSpacing
Horizontal spacing in inches for grid/stack layout.

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

### -HorizontalStack
Lay out selected shapes as a horizontal stack.

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

### -InputObject
Shapes, shape selections, or shape keys/ids to arrange.

```yaml
Type: Object
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Layer
Add selected shapes to this layer.

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

### -NoRouteInternalConnectors
Do not reroute internal connectors during OfficeIMO relayout.

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

### -Page
Page that owns the shapes. Optional inside OfficeVisio/VisioPage DSL scopes.

```yaml
Type: VisioPage
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
Emit arranged shapes.

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

### -PreserveFirstShapeCenter
Use the first selected shape as the grid origin instead of preserving the selection top-left.

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

### -ShapeId
Shape keys or ids to resolve on the target page.

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

### -VerticalSpacing
Vertical spacing in inches for grid/stack layout.

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

### -VerticalStack
Lay out selected shapes as a vertical stack.

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

- `System.Object`

## OUTPUTS

- `OfficeIMO.Visio.VisioShape`

## RELATED LINKS

- None
