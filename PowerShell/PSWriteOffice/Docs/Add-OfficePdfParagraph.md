---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficePdfParagraph
## SYNOPSIS
Adds a plain or rich PDF paragraph using PdfText children.

## SYNTAX
### Text (Default)
```powershell
Add-OfficePdfParagraph [-Text] <string> [-Alignment <OfficeAlignment>] [-Color <OfficeColor>] [<CommonParameters>]
```

### DocumentText
```powershell
Add-OfficePdfParagraph [-Text] <string> -Document <PdfDocument> [-Alignment <OfficeAlignment>] [-Color <OfficeColor>] [-PassThru] [<CommonParameters>]
```

### DocumentContent
```powershell
Add-OfficePdfParagraph [-Content] <scriptblock> -Document <PdfDocument> [-Alignment <OfficeAlignment>] [-Color <OfficeColor>] [-PassThru] [<CommonParameters>]
```

### Content
```powershell
Add-OfficePdfParagraph [-Content] <scriptblock> [-Alignment <OfficeAlignment>] [-Color <OfficeColor>] [<CommonParameters>]
```

## DESCRIPTION
Adds a plain or rich PDF paragraph using PdfText children.

## EXAMPLES

### EXAMPLE 1
```powershell
New-OfficePdf ./Report.pdf { PdfParagraph { PdfText 'Status: '; PdfText 'Healthy' -Bold } }
```


## PARAMETERS

### -Alignment
Paragraph alignment.

```yaml
Type: OfficeAlignment
Parameter Sets: Text, DocumentText, DocumentContent, Content
Aliases: None
Possible values: Left, Center, Right, Justify, Start, End

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Color
Default paragraph color.

```yaml
Type: OfficeColor
Parameter Sets: Text, DocumentText, DocumentContent, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Content
Rich paragraph composed with PdfText.

```yaml
Type: ScriptBlock
Parameter Sets: DocumentContent, Content
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Live PDF outside a composition block.

```yaml
Type: PdfDocument
Parameter Sets: DocumentText, DocumentContent
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -PassThru
Emit the updated live document.

```yaml
Type: SwitchParameter
Parameter Sets: DocumentText, DocumentContent
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Text
Plain paragraph text, including an empty paragraph.

```yaml
Type: String
Parameter Sets: Text, DocumentText
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
