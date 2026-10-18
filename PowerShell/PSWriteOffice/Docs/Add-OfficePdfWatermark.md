---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficePdfWatermark
## SYNOPSIS
Adds a generated-document text watermark.

## SYNTAX
### Context (Default)
```powershell
Add-OfficePdfWatermark [-Text] <string> [-FontSize <Double>] [-Opacity <Double>] [-RotationAngle <Double>] [-Color <string>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficePdfWatermark [-Text] <string> -Document <PdfDocument> [-FontSize <Double>] [-Opacity <Double>] [-RotationAngle <Double>] [-Color <string>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a generated-document text watermark.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficePdf -Path .\Examples\Documents\PdfWatermark.pdf {
    Add-OfficePdfWatermark -Text 'DRAFT' -FontSize 72 -Opacity 0.12 -RotationAngle -35 -Color '#64748B'
    Add-OfficePdfHeading -Text 'Draft service review'
    Add-OfficePdfParagraph -Text 'This copy is not final.'
}
```

Adds a text watermark while generating the PDF.

## PARAMETERS

### -Color
Optional watermark color. Named colors and hexadecimal values are accepted.

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
Compatibility parameter. Page composition is supported only inside New-OfficePdf.

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

### -FontSize
Watermark font size.

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

### -Opacity
Watermark opacity, 0 through 1.

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

### -RotationAngle
Watermark rotation angle.

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

### -Text
Watermark text.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Pdf.PdfDocument`

## OUTPUTS

- `OfficeIMO.Pdf.PdfDocument`

## RELATED LINKS

- None
