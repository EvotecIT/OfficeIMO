---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficePdfText
## SYNOPSIS
Adds a rich inline-text paragraph to a generated PDF document.

## SYNTAX
### Context (Default)
```powershell
Add-OfficePdfText [[-Text] <string[]>] [-Run <Object[]>] [-Align <PdfAlign>] [-Color <string>] [-BackgroundColor <string>] [-FontSize <Double>] [-Font <PdfStandardFont>] [-Bold] [-Italic] [-Underline] [-Strike] [-Baseline <PdfTextBaseline>] [-LinkUri <string>] [-LinkDestinationName <string>] [-LinkContents <string>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficePdfText [[-Text] <string[]>] -Document <PdfDocument> [-Run <Object[]>] [-Align <PdfAlign>] [-Color <string>] [-BackgroundColor <string>] [-FontSize <Double>] [-Font <PdfStandardFont>] [-Bold] [-Italic] [-Underline] [-Strike] [-Baseline <PdfTextBaseline>] [-LinkUri <string>] [-LinkDestinationName <string>] [-LinkContents <string>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Use Add-OfficePdfText when a paragraph needs mixed emphasis, highlight color, font settings, baseline changes, or links.
Plain paragraphs can continue to use Add-OfficePdfParagraph. Rich text runs are translated directly to the OfficeIMO.Pdf paragraph builder.
URI links and bookmark links are supported; a single run cannot target both.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficePdf -Path .\Report.pdf { PdfText 'Approved for review' -Bold -Color '#0F766E' -BackgroundColor '#ECFDF5' }
```

Creates a PDF with one styled paragraph.

### EXAMPLE 2
```powershell
PS> New-OfficePdf -Path .\Report.pdf {
  PdfBookmark 'summary'
  PdfText -Run @(
    @{ Text = 'Read the ' }
    @{ Text = 'website'; LinkUri = 'https://evotec.xyz'; Color = '#2563EB' }
    @{ Text = ' or jump to ' }
    @{ Text = 'summary'; LinkDestinationName = 'summary'; Color = '#7C3AED' }
    @{ Text = '.' }
  )
}
```

Creates one paragraph with an external link and an internal named-destination link.

## PARAMETERS

### -Align
Paragraph alignment.

```yaml
Type: PdfAlign
Parameter Sets: Context, Document
Aliases: None
Possible values: Left, Center, Right, Justify

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -BackgroundColor
Run background color for -Text input. Named colors and hexadecimal colors are accepted.

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

### -Baseline
Baseline for -Text input.

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
Make -Text input bold.

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
Default paragraph color. Named colors and hexadecimal colors are accepted.

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

### -Document
PDF document to update outside the DSL context.

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
Standard PDF font for -Text input.

```yaml
Type: PdfStandardFont
Parameter Sets: Context, Document
Aliases: None
Possible values: Helvetica, HelveticaOblique, HelveticaBold, HelveticaBoldOblique, TimesRoman, TimesItalic, TimesBold, TimesBoldItalic, Courier, CourierOblique, CourierBold, CourierBoldOblique

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FontSize
Font size for -Text input in PDF points.

```yaml
Type: Double
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
Make -Text input italic.

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

### -LinkContents
Optional link annotation contents for -Text input.

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

### -LinkDestinationName
Named destination link target for -Text input.

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

### -LinkUri
Absolute URI link target for -Text input.

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
Emit the updated document.

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

### -Run
Rich run specifications. Each run may define Text, Bold, Italic, Underline, Strike, Color, BackgroundColor, FontSize, Font, Baseline, LinkUri, LinkDestinationName, LinkContents, Type, or Kind.

```yaml
Type: Object[]
Parameter Sets: Context, Document
Aliases: Runs
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Strike
Strike through -Text input.

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

### -Text
Plain text values to add as one styled paragraph.

```yaml
Type: String[]
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Underline
Underline -Text input.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Pdf.PdfDocument`

## OUTPUTS

- `OfficeIMO.Pdf.PdfDocument`

## RELATED LINKS

- None
