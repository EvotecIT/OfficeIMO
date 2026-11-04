---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficePdfText
## SYNOPSIS
Adds one styled run inside a PDF paragraph or running header/footer, or a standalone paragraph.

## SYNTAX
### Context (Default)
```powershell
Add-OfficePdfText [-Text] <string> [-Bold] [-Italic] [-Strikethrough] [-Underline <OfficeTextDecorationStyle>] [-Color <OfficeColor>] [-BackgroundColor <OfficeColor>] [-Font <string>] [-Size <OfficeLength>] [-Baseline <PdfTextBaseline>] [-TabLeader <PdfTabLeaderStyle>] [-TabAlignment <PdfTabAlignment>] [-Link <string>] [-LinkBookmark <string>] [-LinkContents <string>] [-Alignment <OfficeAlignment>] [<CommonParameters>]
```

### Document
```powershell
Add-OfficePdfText [-Text] <string> -Document <PdfDocument> [-Bold] [-Italic] [-Strikethrough] [-Underline <OfficeTextDecorationStyle>] [-Color <OfficeColor>] [-BackgroundColor <OfficeColor>] [-Font <string>] [-Size <OfficeLength>] [-Baseline <PdfTextBaseline>] [-TabLeader <PdfTabLeaderStyle>] [-TabAlignment <PdfTabAlignment>] [-Link <string>] [-LinkBookmark <string>] [-LinkContents <string>] [-Alignment <OfficeAlignment>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds one styled run inside a PDF paragraph or running header/footer, or a standalone paragraph.

## EXAMPLES

### EXAMPLE 1
```powershell
PdfParagraph { PdfText 'Review ' -Italic; PdfText 'complete' -Bold -Color Success }
```


## PARAMETERS

### -Alignment
Standalone paragraph alignment.

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

### -BackgroundColor
Run background color.

```yaml
Type: OfficeColor
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Baseline
Run baseline.

```yaml
Type: PdfTextBaseline
Parameter Sets: Context, Document
Aliases: None
Possible values: Normal, Superscript, Subscript

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Bold
Set bold on or off.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Color
Run foreground color.

```yaml
Type: OfficeColor
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Live PDF outside a composition block.

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

### -Font
Standard PDF font name or registered embedded family.

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

### -Italic
Set italic on or off.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Link
URI hyperlink target.

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

### -LinkBookmark
Named bookmark target.

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

### -LinkContents
Link annotation contents.

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

### -PassThru
Emit the updated live document.

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

### -Size
Font size; bare numbers mean points.

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

### -Strikethrough
Set strikethrough on or off.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TabAlignment
Alignment for an explicit tab text.

```yaml
Type: PdfTabAlignment
Parameter Sets: Context, Document
Aliases: None
Possible values: Left, Right, Center, DecimalSeparator, Bar

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TabLeader
Leader pattern for an explicit tab text (a single tab character).

```yaml
Type: PdfTabLeaderStyle
Parameter Sets: Context, Document
Aliases: None
Possible values: None, Dots, Hyphens, Underscores

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Text
Text for this run.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Underline
Underline pattern; None clears it.

```yaml
Type: OfficeTextDecorationStyle
Parameter Sets: Context, Document
Aliases: None
Possible values: None, Single, Double, Dotted, Dashed, Wavy, Words

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Pdf.PdfDocument`

## OUTPUTS

- `OfficeIMO.Pdf.PdfDocument`

## RELATED LINKS

- None
