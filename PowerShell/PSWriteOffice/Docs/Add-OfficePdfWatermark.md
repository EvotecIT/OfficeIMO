---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficePdfWatermark
## SYNOPSIS
Adds a generated-document text watermark.

## SYNTAX
### Context (Default)
```powershell
Add-OfficePdfWatermark [-Text] <string> [-Size <OfficeLength>] [-Opacity <Double>] [-Rotation <Double>] [-Color <OfficeColor>] [<CommonParameters>]
```

## DESCRIPTION
Adds a generated-document text watermark.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficePdf -Path .\Examples\Documents\PdfWatermark.pdf {
    Add-OfficePdfWatermark -Text 'DRAFT' -Size 72 -Opacity 0.12 -Rotation -35 -Color '#64748B'
    Add-OfficePdfHeading -Text 'Draft service review'
    Add-OfficePdfParagraph -Text 'This copy is not final.'
}
```

Adds a text watermark while generating the PDF.

## PARAMETERS

### -Color
Optional watermark color. Named colors and hexadecimal values are accepted.

```yaml
Type: OfficeColor
Parameter Sets: Context
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
Parameter Sets: Context
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Rotation
Watermark rotation angle.

```yaml
Type: Double
Parameter Sets: Context
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Size
Watermark font size.

```yaml
Type: OfficeLength
Parameter Sets: Context
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
Parameter Sets: Context
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

- `None`

## OUTPUTS

- `None`

## RELATED LINKS

- None
