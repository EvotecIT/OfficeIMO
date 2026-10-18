---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# New-OfficeWordOpenDocumentOptions
## SYNOPSIS
Creates Word/OpenDocument conversion settings.

## SYNTAX
### __AllParameterSets
```powershell
New-OfficeWordOpenDocumentOptions [-LossPolicy <OdfConversionLossPolicy>] [-IncludeImages] [-IncludeHeadersAndFooters] [<CommonParameters>]
```

## DESCRIPTION
Creates Word/OpenDocument conversion settings.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $options = New-OfficeWordOpenDocumentOptions -IncludeImages -IncludeHeadersAndFooters
ConvertTo-OfficeOpenDocument -Path .\Report.docx -OutputPath .\Report.odt -WordOptions $options
```


## PARAMETERS

### -IncludeHeadersAndFooters
Copy default headers and footers.

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

### -IncludeImages
Copy supported inline images.

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

### -LossPolicy
Whether conversion loss is reported or rejected.

```yaml
Type: OdfConversionLossPolicy
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: ReportOnly, ThrowOnSkippedOrUnsupported, ThrowOnAnyLoss

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

- `OfficeIMO.Word.OpenDocument.WordOpenDocumentConversionOptions`

## RELATED LINKS

- None
