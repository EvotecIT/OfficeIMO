---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficePdfParagraph
## SYNOPSIS
Adds a paragraph to a PDF document.

## SYNTAX
### Context (Default)
```powershell
Add-OfficePdfParagraph [-Text] <string> [-Align <PdfAlign>] [-Color <string>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficePdfParagraph [-Text] <string> -Document <PdfDocument> [-Align <PdfAlign>] [-Color <string>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a paragraph to a PDF document.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficePdf -Path .\Examples\Documents\PdfParagraph.pdf {
    Add-OfficePdfHeading -Text 'Status'
    Add-OfficePdfParagraph -Text 'All monitored services are currently healthy.' -Color '#166534'
}
```

Adds a colored body paragraph after a heading.

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

### -Color
Optional text color. Named colors and hexadecimal values are accepted.

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

### -Text
Paragraph text.

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
