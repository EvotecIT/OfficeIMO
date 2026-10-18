---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficePdfCanvasText
## SYNOPSIS
Adds PowerShell-friendly text or rich text runs to the active fixed-position PDF canvas.

## SYNTAX
### Text (Default)
```powershell
Add-OfficePdfCanvasText [-Text] <string[]> -X <double> -Y <double> [-Width <Double>] [-Height <Double>] [-Color <string>] [-Align <PdfAlign>] [-FontSize <Double>] [-LineHeight <Double>] [-Bold] [-Italic] [-Underline] [-Strike] [-BackgroundColor <string>] [-Font <PdfStandardFont>] [-Baseline <PdfTextBaseline>] [-PassThru] [<CommonParameters>]
```

### Run
```powershell
Add-OfficePdfCanvasText -Run <Object[]> -X <double> -Y <double> [-Width <Double>] [-Height <Double>] [-Color <string>] [-Align <PdfAlign>] [-FontSize <Double>] [-LineHeight <Double>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Use this command inside Add-OfficePdfCanvas -Content. Coordinates are PDF points measured
from the visual top-left of the page. Rich runs accept TextRun output, hashtables, and objects;
callers do not need to construct native PdfTextRun arrays. Fixed-position canvas
runs are visual text and do not support link targets. Width and height default to the remaining
page area from the supplied coordinates.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Add-OfficePdfCanvas -Path .\Report.pdf -OutputPath .\Stamped.pdf -Content {
    PdfCanvasText -Run @(
      TextRun 'Owner: ' -Bold
      TextRun 'Platform' -Color '#0F766E'
    ) -X 36 -Y 24 -FontSize 10
}
```

The enclosing callback supplies the active page, while the run collection remains an ordinary PowerShell array.

## PARAMETERS

### -Align
Text alignment within the positioned rectangle.

```yaml
Type: PdfAlign
Parameter Sets: Text, Run
Aliases: None
Possible values: Left, Center, Right, Justify

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -BackgroundColor
Background color for plain -Text input.

```yaml
Type: String
Parameter Sets: Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Baseline
Baseline for plain -Text input.

```yaml
Type: PdfTextBaseline
Parameter Sets: Text
Aliases: None
Possible values: Normal, Superscript, Subscript

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Bold
Make plain -Text input bold.

```yaml
Type: SwitchParameter
Parameter Sets: Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Color
Default text color. Named and hexadecimal colors are accepted.

```yaml
Type: String
Parameter Sets: Text, Run
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Font
Standard PDF font for plain -Text input.

```yaml
Type: PdfStandardFont
Parameter Sets: Text
Aliases: None
Possible values: Helvetica, HelveticaOblique, HelveticaBold, HelveticaBoldOblique, TimesRoman, TimesItalic, TimesBold, TimesBoldItalic, Courier, CourierOblique, CourierBold, CourierBoldOblique

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FontSize
Default font size in PDF points.

```yaml
Type: Double
Parameter Sets: Text, Run
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Height
Available text height in PDF points. Defaults to the remaining page height.

```yaml
Type: Double
Parameter Sets: Text, Run
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Italic
Make plain -Text input italic.

```yaml
Type: SwitchParameter
Parameter Sets: Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LineHeight
Optional line height in PDF points.

```yaml
Type: Double
Parameter Sets: Text, Run
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
Parameter Sets: Text, Run
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Run
Rich run specifications created with TextRun or supplied as hashtables or objects. Link targets are not supported.

```yaml
Type: Object[]
Parameter Sets: Run
Aliases: Runs
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Strike
Strike through plain -Text input.

```yaml
Type: SwitchParameter
Parameter Sets: Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Text
Plain text values to concatenate into one positioned text block.

```yaml
Type: String[]
Parameter Sets: Text
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Underline
Underline plain -Text input.

```yaml
Type: SwitchParameter
Parameter Sets: Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Width
Available text width in PDF points. Defaults to the remaining page width.

```yaml
Type: Double
Parameter Sets: Text, Run
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -X
Horizontal position in PDF points from the visual left edge.

```yaml
Type: Double
Parameter Sets: Text, Run
Aliases: Left
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Y
Vertical position in PDF points from the visual top edge.

```yaml
Type: Double
Parameter Sets: Text, Run
Aliases: Top
Possible values:

Required: True
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
