---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficePdfPanel
## SYNOPSIS
Adds a visually separated panel paragraph to a PDF document.

## SYNTAX
### Text (Default)
```powershell
Add-OfficePdfPanel [-Text] <string> [-Alignment <OfficeAlignment>] [<CommonParameters>]
```

### Document
```powershell
Add-OfficePdfPanel [-Text] <string> -Document <PdfDocument> [-Alignment <OfficeAlignment>] [-PassThru] [<CommonParameters>]
```

### Content
```powershell
Add-OfficePdfPanel [-Content] <scriptblock> [-Alignment <OfficeAlignment>] [<CommonParameters>]
```

## DESCRIPTION
Adds a visually separated panel paragraph to a PDF document.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficePdf -Path .\Examples\Documents\PdfPanel.pdf {
    Add-OfficePdfHeading -Text 'Executive summary'
    Add-OfficePdfPanel -Text 'No critical incidents were detected in the current reporting window.' -Alignment Center
}
```

Adds a highlighted panel paragraph to the generated PDF.

## PARAMETERS

### -Alignment
Panel alignment.

```yaml
Type: OfficeAlignment
Parameter Sets: Text, Document, Content
Aliases: None
Possible values: Left, Center, Right, Justify, Start, End

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Content
Rich PdfText child content.

```yaml
Type: ScriptBlock
Parameter Sets: Content
Aliases: None
Possible values:

Required: True
Position: 0
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
Parameter Sets: Document
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
Parameter Sets: Text, Document
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
