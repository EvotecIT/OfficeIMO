---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficePdfFooter
## SYNOPSIS
Sets a simple or fully composed running PDF footer.

## SYNTAX
### Context (Default)
```powershell
Set-OfficePdfFooter [[-Text] <string>] [-Compose <scriptblock>] [-Align <PdfAlign>] [-FontSize <Double>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Set-OfficePdfFooter [[-Text] <string>] -Document <PdfDocument> [-Compose <scriptblock>] [-Align <PdfAlign>] [-FontSize <Double>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Sets a simple or fully composed running PDF footer.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficePdf -Path .\Examples\Documents\PdfFooter.pdf {
    Set-OfficePdfFooter -Text 'Page {page} of {pages}' -Align Center -FontSize 8
    Add-OfficePdfHeading -Text 'Report with footer'
    Add-OfficePdfPageBreak
    Add-OfficePdfParagraph -Text 'The footer includes generated page numbers.'
}
```

Uses page placeholders in a running footer.

### EXAMPLE 2
```powershell
PS> New-OfficePdf -Path .\Examples\Documents\RichFooter.pdf {
    Set-OfficePdfFooter -Compose {
        param($footer)
        $label = New-OfficeTextRun -Text 'Confidential - page ' -Bold | ConvertTo-OfficePdfTextRun
        $pageStyle = New-OfficeTextRun -Italic | ConvertTo-OfficePdfTextRun
        $null = $footer.AlignRight().Text({
            param($text)
            $null = $text.Run($label).CurrentPage($pageStyle).Text(' of ').TotalPages($pageStyle)
        })
    }
    Add-OfficePdfParagraph -Text 'Generated report body.'
}
```

Styled page tokens remain live until the document is rendered.

## PARAMETERS

### -Align
Footer alignment.

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

### -Compose
Advanced footer composer. The script receives a PdfFooterBuilder and can configure
default, first-page, and even-page text, zones, images, shapes, rich text, and page tokens.

```yaml
Type: ScriptBlock
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

### -FontSize
Footer font size in PDF points.

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

### -Text
Footer text. Supports {page} and {pages}.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
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
