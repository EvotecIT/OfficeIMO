---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficePdfBackground
## SYNOPSIS
Sets or clears the generated PDF page background color.

## SYNTAX
### Context (Default)
```powershell
Set-OfficePdfBackground [-Color <string>] [-Clear] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Set-OfficePdfBackground -Document <PdfDocument> [-Color <string>] [-Clear] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Sets or clears the generated PDF page background color.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficePdf -Path .\Examples\Documents\PdfBackground.pdf {
    Set-OfficePdfBackground -Color '#F8FAFC'
    Add-OfficePdfHeading -Text 'Report on a soft background'
    Add-OfficePdfParagraph -Text 'The background color applies to generated pages.'
}
```

Applies a page background color before adding content.

## PARAMETERS

### -Clear
Clear the generated PDF page background color.

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
Background color. Named colors and hexadecimal values are accepted.

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
