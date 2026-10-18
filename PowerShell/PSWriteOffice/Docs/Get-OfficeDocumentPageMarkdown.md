---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeDocumentPageMarkdown
## SYNOPSIS
Projects Reader pages into citation-friendly Markdown.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficeDocumentPageMarkdown [-InputObject] <OfficeDocumentReadResult> [-AsString] [-NoPageMarkers] [<CommonParameters>]
```

## DESCRIPTION
Projects Reader pages into citation-friendly Markdown.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Get-OfficeDocument -Path .\Handbook.pdf |
                Get-OfficeDocumentPageMarkdown -AsString |
                Set-Content -Path .\Handbook.pages.md
```

Uses OfficeIMO.Reader page projection and preserves the page provenance in each marker.

## PARAMETERS

### -AsString
Return one combined Markdown string instead of one page result per page.

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

### -InputObject
Normalized document returned by Get-OfficeDocument.

```yaml
Type: OfficeDocumentReadResult
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -NoPageMarkers
Omit portable HTML page markers from the Markdown.

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

- `OfficeIMO.Reader.OfficeDocumentReadResult`

## OUTPUTS

- `OfficeIMO.Reader.OfficeDocumentPageMarkdown`
- `System.String`

## RELATED LINKS

- None
