---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficePdfPanel
## SYNOPSIS
Adds a visually separated panel paragraph to a PDF document.

## SYNTAX
### Context (Default)
```powershell
Add-OfficePdfPanel [-Text] <string> [-Align <PdfAlign>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficePdfPanel [-Text] <string> -Document <PdfDocument> [-Align <PdfAlign>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a visually separated panel paragraph to a PDF document.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficePdf -Path .\Examples\Documents\PdfPanel.pdf {
    Add-OfficePdfHeading -Text 'Executive summary'
    Add-OfficePdfPanel -Text 'No critical incidents were detected in the current reporting window.' -Align Center
}
```

Adds a highlighted panel paragraph to the generated PDF.

## PARAMETERS

### -Align
Panel alignment.

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
Panel text.

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
