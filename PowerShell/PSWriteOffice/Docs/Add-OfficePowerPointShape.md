---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficePowerPointShape
## SYNOPSIS
Adds a basic shape to a slide.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficePowerPointShape [-Slide <PowerPointSlide>] [-ShapeType <OfficePresetShapeType>] [-X <double>] [-Y <double>] [-Width <double>] [-Height <double>] [-Name <string>] [-FillColor <string>] [-OutlineColor <string>] [-OutlineWidth <Double>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Creates an auto shape at the requested coordinates and applies optional fill and outline styling.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficePowerPoint -Path .\Examples\Documents\PowerPointShape.pptx {
    $slide = Add-OfficePowerPointSlide -Layout 1 -PassThru
    Add-OfficePowerPointShape -Slide $slide -ShapeType Rectangle -X 60 -Y 120 -Width 220 -Height 90 -FillColor '#DDEEFF' -OutlineColor '#2563EB' -OutlineWidth 1
    Add-OfficePowerPointTextBox -Slide $slide -Text 'Highlighted status' -X 80 -Y 145 -Width 180 -Height 32
}
```

Creates a styled rectangle and overlays a text box.

## PARAMETERS

### -FillColor
Fill color (hex or named color).

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

### -Height
Shape height in points.

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

### -Name
Optional name assigned to the shape.

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
Outline color (hex or named color).

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

### -OutlineWidth
Outline width in points.

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

### -PassThru
Emit the object created or changed by the command.

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

### -ShapeType
Shape geometry preset name (e.g., Rectangle, Ellipse, Line).

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

### -Slide
Target slide that will receive the shape (optional inside DSL).

```yaml
Type: PowerPointSlide
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Width
Shape width in points.

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

### -X
Left offset (in points) from the slide origin.

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

### -Y
Top offset (in points) from the slide origin.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.PowerPoint.PowerPointSlide`

## OUTPUTS

- `None`

## RELATED LINKS

- None
