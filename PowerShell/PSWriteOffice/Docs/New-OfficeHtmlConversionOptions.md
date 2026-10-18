---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# New-OfficeHtmlConversionOptions
## SYNOPSIS
Creates discoverable parsing, trust, and document settings for HTML conversion.

## SYNTAX
### __AllParameterSets
```powershell
New-OfficeHtmlConversionOptions [-Profile <HtmlConversionProfile>] [-Trust <HtmlInputTrust>] [-BaseUri <string>] [-UseBodyContentsOnly] [-IncludeNormalizedHtml] [<CommonParameters>]
```

## DESCRIPTION
Creates discoverable parsing, trust, and document settings for HTML conversion.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $document = New-OfficeHtmlConversionOptions -BaseUri (Resolve-Path .\Assets) -UseBodyContentsOnly
Export-OfficeHtmlImage -Path .\Report.html -OutputPath .\Report.svg -DocumentOptions $document
```


## PARAMETERS

### -BaseUri
Base URI used to resolve relative references.

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

### -IncludeNormalizedHtml
Retain normalized HTML in the conversion document.

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

### -Profile
Built-in conversion profile.

```yaml
Type: HtmlConversionProfile
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Semantic, Document, HighFidelityPrint, PositionedReview

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Trust
Input trust level.

```yaml
Type: HtmlInputTrust
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Untrusted, Trusted

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -UseBodyContentsOnly
Convert only body contents.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `OfficeIMO.Html.HtmlConversionDocumentOptions`

## RELATED LINKS

- None
