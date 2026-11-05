---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficeHtmlPdfOptions
## SYNOPSIS
Creates HTML-to-PDF options for Convert-OfficeDocument -To Pdf.

## SYNTAX
### Default (Default)
```powershell
New-OfficeHtmlPdfOptions [-BaseOptions <HtmlPdfExportOptions>] [-PdfOptions <PdfCreateOptions>] [-Font <string>] [-Margin <OfficeLength>] [-IncludeLocalResources] [-NoInteractiveFormControls] [<CommonParameters>]
```

## DESCRIPTION
Creates HTML-to-PDF options for Convert-OfficeDocument -To Pdf.

## EXAMPLES

### EXAMPLE 1
```powershell
Convert-OfficeDocument -Path .\Report.html -To Pdf -OutputPath .\Report.pdf -Options (New-OfficeHtmlPdfOptions -Margin 2cm)
```


## PARAMETERS

### -BaseOptions
BaseOptions for this operation.

```yaml
Type: HtmlPdfExportOptions
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Font
Font for this operation.

```yaml
Type: String
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IncludeLocalResources
IncludeLocalResources for this operation.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Margin
Margin for this operation.

```yaml
Type: OfficeLength
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoInteractiveFormControls
NoInteractiveFormControls for this operation.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PdfOptions
PdfOptions for this operation.

```yaml
Type: PdfCreateOptions
Parameter Sets: Default
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

- `OfficeIMO.Html.Pdf.HtmlPdfExportOptions`

## OUTPUTS

- `OfficeIMO.Html.Pdf.HtmlPdfExportOptions`

## RELATED LINKS

- None
