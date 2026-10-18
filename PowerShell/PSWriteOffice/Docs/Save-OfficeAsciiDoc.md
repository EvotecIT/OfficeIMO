---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Save-OfficeAsciiDoc
## SYNOPSIS
Saves an OfficeIMO AsciiDoc document.

## SYNTAX
### __AllParameterSets
```powershell
Save-OfficeAsciiDoc [-Document] <AsciiDocDocument> [-Path] <string> [-Options <AsciiDocWriterOptions>] [-Mode <AsciiDocWriterMode>] [-LineEnding <OfficeLineEnding>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Saves an OfficeIMO AsciiDoc document.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $document = Get-OfficeAsciiDoc -Path .\Guide.adoc
$document | Save-OfficeAsciiDoc -Path .\Guide-normalized.adoc -Mode Canonical
```


## PARAMETERS

### -Document
AsciiDoc document to save.

```yaml
Type: AsciiDocDocument
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -LineEnding
Canonical line ending: LF, CRLF, or CR. Omit it to retain the source preference.

```yaml
Type: OfficeLineEnding
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: LF, CRLF, CR

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Mode
Writer mode. Preserve retains unchanged source; Canonical emits stable formatting.

```yaml
Type: AsciiDocWriterMode
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Preserve, Canonical

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Options
Optional writer settings.

```yaml
Type: AsciiDocWriterOptions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Return the saved document.

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

### -Path
Destination path.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.AsciiDoc.AsciiDocDocument`

## OUTPUTS

- `OfficeIMO.AsciiDoc.AsciiDocDocument`

## RELATED LINKS

- None
