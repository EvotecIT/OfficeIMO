---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficePdfHeader
## SYNOPSIS
Sets a simple or fully composed running PDF header.

## SYNTAX
### Context (Default)
```powershell
Set-OfficePdfHeader [[-Text] <string>] [-Compose <scriptblock>] [-Align <PdfAlign>] [-FontSize <Double>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Set-OfficePdfHeader [[-Text] <string>] -Document <PdfDocument> [-Compose <scriptblock>] [-Align <PdfAlign>] [-FontSize <Double>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Sets a simple or fully composed running PDF header.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficePdf -Path .\Examples\Documents\PdfHeader.pdf {
    Set-OfficePdfHeader -Text 'Service Review' -Align Right -FontSize 9
    Add-OfficePdfHeading -Text 'Service Review'
    Add-OfficePdfParagraph -Text 'The header repeats on generated pages.'
}
```

Sets header text for the generated PDF.

### EXAMPLE 2
```powershell
PS> New-OfficePdf -Path .\Examples\Documents\RichHeader.pdf {
    Set-OfficePdfHeader -Compose {
        param($header)
        $label = New-OfficeTextRun -Text 'Service report ' -Bold | ConvertTo-OfficePdfTextRun
        $pageStyle = New-OfficeTextRun -Italic | ConvertTo-OfficePdfTextRun
        $null = $header.Text({
            param($text)
            $null = $text.Run($label).CurrentPage($pageStyle)
        })
        $null = $header.FirstPageText('Service report cover')
        $null = $header.EvenPagesZones('Service report', $null, 'Page {page}/{pages}')
    }
    Add-OfficePdfParagraph -Text 'Generated report body.'
}
```

The native composer owns rich runs, page tokens, zones, images, shapes, and page variants.

## PARAMETERS

### -Align
Header alignment.

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
Advanced header composer. The script receives a PdfHeaderBuilder and can configure
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
Header font size in PDF points.

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
Header text. Supports {page} and {pages}.

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
