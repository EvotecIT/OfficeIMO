---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# New-OfficeHtmlPdfOptions
## SYNOPSIS
Creates typed HTML rendering options for Export-OfficeDocumentPdf.

## SYNTAX
### __AllParameterSets
```powershell
New-OfficeHtmlPdfOptions [-Options <HtmlToPdfOptions>] [-PdfOptions <PdfOptions>] [-FontFamily <string>] [-Margin <Double>] [-IncludeLocalResources] [-InteractiveFormControls <Boolean>] [<CommonParameters>]
```

## DESCRIPTION
Creates typed HTML rendering options for Export-OfficeDocumentPdf.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $options = New-OfficeHtmlPdfOptions -FontFamily Arial -IncludeLocalResources
Export-OfficeDocumentPdf -InputPath .\Report.html -Path .\Report.pdf -HtmlOptions $options
```


## PARAMETERS

### -FontFamily
Default HTML font family.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IncludeLocalResources
Allow bounded local images, stylesheets and fonts. Workflow batches constrain them to the source directory.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -InteractiveFormControls
Render supported HTML form controls as interactive PDF fields.

```yaml
Type: Boolean
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Margin
Uniform page margins in CSS pixels.

```yaml
Type: Double
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Options
Existing options to clone before applying explicitly supplied values.

```yaml
Type: HtmlToPdfOptions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -PdfOptions
Underlying PDF writer settings, cloned before use.

```yaml
Type: PdfOptions
Parameter Sets: __AllParameterSets
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

- `OfficeIMO.Html.Pdf.HtmlToPdfOptions`

## OUTPUTS

- `OfficeIMO.Html.Pdf.HtmlToPdfOptions`

## RELATED LINKS

- None
