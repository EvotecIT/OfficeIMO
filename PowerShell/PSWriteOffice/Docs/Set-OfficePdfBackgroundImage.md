---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficePdfBackgroundImage
## SYNOPSIS
Sets or clears a generated PDF page background image.

## SYNTAX
### Context (Default)
```powershell
Set-OfficePdfBackgroundImage [[-Path] <string>] [-Fit <OfficeImageFit>] [-Opacity <Double>] [-Clear] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Set-OfficePdfBackgroundImage [[-Path] <string>] -Document <PdfDocument> [-Fit <OfficeImageFit>] [-Opacity <Double>] [-Clear] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Background images are applied through OfficeIMO.Pdf and rendered behind normal generated content.
Use low opacity for watermark-like page texture and -Clear to remove a previously configured background image.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficePdf -Path .\Report.pdf {
  PdfBackgroundImage -Path .\letterhead.png -Fit Cover -Opacity 0.08
  PdfHeading 'Branded report'
}
```

Uses an image as a low-opacity page background.

## PARAMETERS

### -Clear
Clear the generated PDF background image.

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

### -Fit
How the image should fit the page box.

```yaml
Type: OfficeImageFit
Parameter Sets: Context, Document
Aliases: None
Possible values: Stretch, Contain, Cover

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Opacity
Image opacity from 0 to 1.

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

### -Path
Background image path.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: FilePath
Possible values:

Required: False
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
