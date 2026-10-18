---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficePdfPageBreak
## SYNOPSIS
Adds a page break to a PDF document.

## SYNTAX
### Context (Default)
```powershell
Add-OfficePdfPageBreak [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficePdfPageBreak -Document <PdfDocument> [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a page break to a PDF document.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficePdf -Path .\Examples\Documents\PdfPageBreak.pdf {
    Add-OfficePdfHeading -Text 'Service review'
    Add-OfficePdfParagraph -Text 'Summary content stays on the first page.'
    Add-OfficePdfPageBreak
    Add-OfficePdfHeading -Text 'Appendix' -Level 2
}
```

Forces the appendix section to begin on the next page.

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
