---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficePdfImage
## SYNOPSIS
Adds an image to a PDF document.

## SYNTAX
### Context (Default)
```powershell
Add-OfficePdfImage [-Path] <string> -Width <double> -Height <double> [-Align <PdfAlign>] [-AlternativeText <string>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficePdfImage [-Path] <string> -Document <PdfDocument> -Width <double> -Height <double> [-Align <PdfAlign>] [-AlternativeText <string>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds an image to a PDF document.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $image = '.\Tests\Assets\CellImage.png'
New-OfficePdf -Path .\Examples\Documents\PdfImage.pdf {
    Add-OfficePdfHeading -Text 'Visual evidence'
    Add-OfficePdfImage -Path $image -Width 120 -Height 80 -Align Center -AlternativeText 'Sample chart image'
}
```

Embeds an image with alternative text in the generated PDF.

## PARAMETERS

### -Align
Image alignment.

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

### -AlternativeText
Alternative text for meaningful images.

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

### -Height
Rendered image height in PDF points.

```yaml
Type: Double
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: True
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

### -Path
Image path.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: FilePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Width
Rendered image width in PDF points.

```yaml
Type: Double
Parameter Sets: Context, Document
Aliases: None
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

- `OfficeIMO.Pdf.PdfDocument`

## OUTPUTS

- `OfficeIMO.Pdf.PdfDocument`

## RELATED LINKS

- None
