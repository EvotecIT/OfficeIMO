---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficePowerPointShape
## SYNOPSIS
Adds a native preset shape at physical coordinates or on the theme grid.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficePowerPointShape [-Slide <PowerPointSlideSelector>] [-Type <OfficePresetShapeType>] [-Name <string>] [-FillColor <OfficeColor>] [-OutlineColor <OfficeColor>] [-OutlineWidth <OfficeLength>] [-X <OfficeLength>] [-Y <OfficeLength>] [-Width <OfficeLength>] [-Height <OfficeLength>] [-Column <Int32>] [-Span <Int32>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a native preset shape at physical coordinates or on the theme grid.

## EXAMPLES

### EXAMPLE 1
```powershell
PowerPointSlide -Title 'Service health' {
    PowerPointShape -Type Ellipse -Name 'Healthy' -FillColor Success -X 2cm -Y 4cm -Width 2cm -Height 2cm
}
```


## PARAMETERS

### -Column
Gets or sets the one-based theme grid column; explicit X overrides its position.

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

### -FillColor
Gets or sets the fill color or theme role.

```yaml
Type: OfficeColor
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Height
Gets or sets the height; omission uses the native default.

```yaml
Type: OfficeLength
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Name
Gets or sets the native shape name.

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

### -OutlineColor
Gets or sets the outline color or theme role.

```yaml
Type: OfficeColor
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutlineWidth
Gets or sets the outline thickness; numbers mean points.

```yaml
Type: OfficeLength
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
Gets or sets whether the live native shape is returned.

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

### -Slide
Gets or sets a native slide or one-based position; omission uses the current slide block.

```yaml
Type: PowerPointSlideSelector
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Span
Gets or sets the number of grid columns; the default is twelve, and explicit Width overrides it.

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

### -Type
Gets or sets the native preset geometry.

```yaml
Type: OfficePresetShapeType
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Line, LineInverse, Triangle, RightTriangle, Rectangle, Diamond, Parallelogram, Trapezoid, NonIsoscelesTrapezoid, Pentagon, Hexagon, Heptagon, Octagon, Decagon, Dodecagon, Star4, Star5, Star6, Star7, Star8, Star10, Star12, Star16, Star24, Star32, RoundRectangle, Round1Rectangle, Round2SameRectangle, Round2DiagonalRectangle, SnipRoundRectangle, Snip1Rectangle, Snip2SameRectangle, Snip2DiagonalRectangle, Plaque, Ellipse, Teardrop, HomePlate, Chevron, PieWedge, Pie, BlockArc, Donut, NoSmoking, RightArrow, LeftArrow, UpArrow, DownArrow, StripedRightArrow, NotchedRightArrow, BentUpArrow, LeftRightArrow, UpDownArrow, LeftUpArrow, LeftRightUpArrow, QuadArrow, LeftArrowCallout, RightArrowCallout, UpArrowCallout, DownArrowCallout, LeftRightArrowCallout, UpDownArrowCallout, QuadArrowCallout, BentArrow, UTurnArrow, CircularArrow, LeftCircularArrow, LeftRightCircularArrow, CurvedRightArrow, CurvedLeftArrow, CurvedUpArrow, CurvedDownArrow, SwooshArrow, Cube, Can, LightningBolt, Heart, Sun, Moon, SmileyFace, IrregularSeal1, IrregularSeal2, FoldedCorner, Bevel, Frame, HalfFrame, Corner, DiagonalStripe, Chord, Arc, LeftBracket, RightBracket, LeftBrace, RightBrace, BracketPair, BracePair, StraightConnector1, BentConnector2, BentConnector3, BentConnector4, BentConnector5, CurvedConnector2, CurvedConnector3, CurvedConnector4, CurvedConnector5, Callout1, Callout2, Callout3, AccentCallout1, AccentCallout2, AccentCallout3, BorderCallout1, BorderCallout2, BorderCallout3, AccentBorderCallout1, AccentBorderCallout2, AccentBorderCallout3, WedgeRectangleCallout, WedgeRoundRectangleCallout, WedgeEllipseCallout, CloudCallout, Cloud, Ribbon, Ribbon2, EllipseRibbon, EllipseRibbon2, LeftRightRibbon, VerticalScroll, HorizontalScroll, Wave, DoubleWave, Plus, FlowChartProcess, FlowChartDecision, FlowChartInputOutput, FlowChartPredefinedProcess, FlowChartInternalStorage, FlowChartDocument, FlowChartMultidocument, FlowChartTerminator, FlowChartPreparation, FlowChartManualInput, FlowChartManualOperation, FlowChartConnector, FlowChartPunchedCard, FlowChartPunchedTape, FlowChartSummingJunction, FlowChartOr, FlowChartCollate, FlowChartSort, FlowChartExtract, FlowChartMerge, FlowChartOfflineStorage, FlowChartOnlineStorage, FlowChartMagneticTape, FlowChartMagneticDisk, FlowChartMagneticDrum, FlowChartDisplay, FlowChartDelay, FlowChartAlternateProcess, FlowChartOffpageConnector, ActionButtonBlank, ActionButtonHome, ActionButtonHelp, ActionButtonInformation, ActionButtonForwardNext, ActionButtonBackPrevious, ActionButtonEnd, ActionButtonBeginning, ActionButtonReturn, ActionButtonDocument, ActionButtonSound, ActionButtonMovie, Gear6, Gear9, Funnel, MathPlus, MathMinus, MathMultiply, MathDivide, MathEqual, MathNotEqual, CornerTabs, SquareTabs, PlaqueTabs, ChartX, ChartStar, ChartPlus

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Width
Gets or sets the width; omission uses the native default.

```yaml
Type: OfficeLength
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -X
Gets or sets the left coordinate; numbers mean points.

```yaml
Type: OfficeLength
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Y
Gets or sets the top coordinate; numbers mean points.

```yaml
Type: OfficeLength
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

- `PSWriteOffice.Foundation.PowerPointSlideSelector`: A live slide or a one-based slide position in a document.

## OUTPUTS

- `OfficeIMO.PowerPoint.PowerPointAutoShape`

## RELATED LINKS

- None
