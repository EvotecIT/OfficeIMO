---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeAsciiDoc
## SYNOPSIS
Parses an AsciiDoc file or source string into OfficeIMO's native document model.

## SYNTAX
### Path (Default)
```powershell
Get-OfficeAsciiDoc [-Path] <string> [-Options <AsciiDocParseOptions>] [-AsResult] [<CommonParameters>]
```

### Text
```powershell
Get-OfficeAsciiDoc -Text <string> [-Options <AsciiDocParseOptions>] [-AsResult] [<CommonParameters>]
```

## DESCRIPTION
Parses an AsciiDoc file or source string into OfficeIMO's native document model.

## EXAMPLES

### EXAMPLE 1
```powershell
Get-OfficeAsciiDoc -Path 'C:\Path'
```


### EXAMPLE 2
```powershell
Get-OfficeAsciiDoc -Text 'Value'
```


## PARAMETERS

### -AsResult
Return the parse result with diagnostics instead of only the document.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Options
Optional parser settings.

```yaml
Type: AsciiDocParseOptions
Parameter Sets: Path, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Path to an AsciiDoc file.

```yaml
Type: String
Parameter Sets: Path
Aliases: FilePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Text
AsciiDoc source text.

```yaml
Type: String
Parameter Sets: Text
Aliases: None
Possible values:

Required: True
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

- `OfficeIMO.AsciiDoc.AsciiDocDocument`
- `OfficeIMO.AsciiDoc.AsciiDocParseResult`

## RELATED LINKS

- None
