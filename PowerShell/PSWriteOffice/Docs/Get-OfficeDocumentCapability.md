---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeDocumentCapability
## SYNOPSIS
Lists OfficeIMO.Reader capabilities registered in the current PSWriteOffice process.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficeDocumentCapability [-Manifest] [-ExcludeBuiltIn] [-ExcludeCustom] [-Reader <OfficeDocumentReader>] [<CommonParameters>]
```

## DESCRIPTION
Lists OfficeIMO.Reader capabilities registered in the current PSWriteOffice process.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $capabilities = Get-OfficeDocumentCapability
$capabilities | Sort-Object Id | Select-Object Id, Extensions
```

Lists built-in and modular Reader handlers, including adapters such as PDF, RTF, HTML, CSV, JSON, XML, YAML, ZIP, EPUB, and Visio when available.

## PARAMETERS

### -ExcludeBuiltIn
Exclude built-in Reader capabilities.

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

### -ExcludeCustom
Exclude custom or modular Reader capabilities.

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

### -Manifest
Return the capability manifest envelope instead of individual handlers.

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

### -Reader
Optional immutable OfficeIMO reader with caller-configured handlers and processors.

```yaml
Type: OfficeDocumentReader
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

- `OfficeIMO.Reader.ReaderHandlerCapability`
- `OfficeIMO.Reader.ReaderCapabilityManifest`

## RELATED LINKS

- None
