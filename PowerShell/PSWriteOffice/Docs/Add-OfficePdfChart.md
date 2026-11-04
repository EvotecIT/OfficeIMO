---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficePdfChart
## SYNOPSIS
Adds a themed chart from data.

## SYNTAX
### Context (Default)
```powershell
Add-OfficePdfChart -InputObject <Object[]> -Category <string> -Value <string[]> [-Type <OfficeChartType>] [-Title <string>] [-Width <OfficeLength>] [-Height <OfficeLength>] [-Alignment <OfficeAlignment>] [-AlternativeText <string>] [<CommonParameters>]
```

### Document
```powershell
Add-OfficePdfChart -InputObject <Object[]> -Category <string> -Value <string[]> -Document <PdfDocument> [-Type <OfficeChartType>] [-Title <string>] [-Width <OfficeLength>] [-Height <OfficeLength>] [-Alignment <OfficeAlignment>] [-AlternativeText <string>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a themed chart from data.

## EXAMPLES

### EXAMPLE 1
```powershell
$sales = @(
    [pscustomobject]@{ Region = 'North'; Q1 = 1.42; Q2 = 1.55; Q3 = 1.71 }
    [pscustomobject]@{ Region = 'South'; Q1 = 0.98; Q2 = 1.04; Q3 = 1.12 }
    [pscustomobject]@{ Region = 'East'; Q1 = 1.21; Q2 = 1.18; Q3 = 1.36 }
    [pscustomobject]@{ Region = 'West'; Q1 = 1.65; Q2 = 1.79; Q3 = 1.88 }
)
New-OfficePdf -Path .\Sales.pdf {
    PdfChart -InputObject $sales -Category Region -Value Q3 -Title 'Q3 revenue by region (USD m)'
}
```


## PARAMETERS

### -Alignment
Alignment for this operation.

```yaml
Type: OfficeAlignment
Parameter Sets: Context, Document
Aliases: None
Possible values: Left, Center, Right, Justify, Start, End

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AlternativeText
AlternativeText for this operation.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Category
Category for this operation.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Document for this operation.

```yaml
Type: PdfDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Height
Height for this operation.

```yaml
Type: OfficeLength
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -InputObject
InputObject for this operation.

```yaml
Type: Object[]
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -PassThru
Document set only; inside a block the document is returned by New-OfficePdf.

```yaml
Type: SwitchParameter
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Title
Title for this operation.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Type
Type for this operation.

```yaml
Type: OfficeChartType
Parameter Sets: Context, Document
Aliases: None
Possible values: ColumnClustered, Column, ColumnStacked, ColumnStacked100, BarClustered, Bar, BarStacked, BarStacked100, Line, LineStacked, LineStacked100, Area, AreaStacked, AreaStacked100, Scatter, Radar, Pie, Doughnut, Bubble, Column3DClustered, Column3DStacked, Column3DStacked100, Bar3DClustered, Bar3DStacked, Bar3DStacked100, Line3D, Area3D, Area3DStacked, Area3DStacked100, Pie3D, PieOfPie, BarOfPie, Stock, Surface, SurfaceWireframe, SurfaceContour, Funnel, Waterfall, BoxWhisker, Treemap, SurfaceContourWireframe, Sunburst

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Value
Value for this operation.

```yaml
Type: String[]
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Width
Width for this operation.

```yaml
Type: OfficeLength
Parameter Sets: Context, Document
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
- `OfficeIMO.Pdf.PdfDocument`

## OUTPUTS

- `OfficeIMO.Pdf.PdfDocument`

## RELATED LINKS

- None
