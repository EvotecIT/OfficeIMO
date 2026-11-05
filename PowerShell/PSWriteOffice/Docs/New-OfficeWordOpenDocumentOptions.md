---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficeWordOpenDocumentOptions
## SYNOPSIS
Creates Word/OpenDocument conversion settings.

## SYNTAX
### Default (Default)
```powershell
New-OfficeWordOpenDocumentOptions [-LossPolicy <OdfConversionLossPolicy>] [-NoImages] [-NoHeadersAndFooters] [<CommonParameters>]
```

## DESCRIPTION
Creates Word/OpenDocument conversion settings.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $options = New-OfficeWordOpenDocumentOptions -NoImages -NoHeadersAndFooters
Convert-OfficeDocument -Path .\Report.docx -To OpenDocument -OutputPath .\Report.odt -Options $options
```


## PARAMETERS

### -LossPolicy
Whether conversion loss is reported or rejected.

```yaml
Type: OdfConversionLossPolicy
Parameter Sets: Default
Aliases: None
Possible values: ReportOnly, ThrowOnSkippedOrUnsupported, ThrowOnAnyLoss

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoHeadersAndFooters
Copy default headers and footers.

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

### -NoImages
Copy supported inline images.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `OfficeIMO.Word.OpenDocument.WordOpenDocumentConversionOptions`

## RELATED LINKS

- None
