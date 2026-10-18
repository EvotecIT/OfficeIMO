---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficePdfSpacer
## SYNOPSIS
Adds invisible vertical spacing to a generated PDF document.

## SYNTAX
### Context (Default)
```powershell
Add-OfficePdfSpacer [-Height] <double> [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficePdfSpacer [-Height] <double> -Document <PdfDocument> [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds invisible vertical spacing to a generated PDF document.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficePdf -Path .\Examples\Documents\PdfSpacer.pdf {
    Add-OfficePdfHeading -Text 'Summary'
    Add-OfficePdfParagraph -Text 'First block.'
    Add-OfficePdfSpacer -Height 18
    Add-OfficePdfParagraph -Text 'Second block after additional spacing.'
}
```

Adds whitespace without adding visible content.

## PARAMETERS

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
Vertical space height in PDF points.

```yaml
Type: Double
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: True
Position: 0
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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Pdf.PdfDocument`

## OUTPUTS

- `OfficeIMO.Pdf.PdfDocument`

## RELATED LINKS

- None
