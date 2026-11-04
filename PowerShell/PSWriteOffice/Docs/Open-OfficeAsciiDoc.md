---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Open-OfficeAsciiDoc
## SYNOPSIS
Parses an AsciiDoc file or source string into OfficeIMO's native document model.

## SYNTAX
### Path (Default)
```powershell
Open-OfficeAsciiDoc [-Path] <string> [-Options <AsciiDocParseOptions>] [-AsResult] [<CommonParameters>]
```

### LiteralPath
```powershell
Open-OfficeAsciiDoc -LiteralPath <string> [-Options <AsciiDocParseOptions>] [-AsResult] [<CommonParameters>]
```

### Text
```powershell
Open-OfficeAsciiDoc -Text <string> [-Options <AsciiDocParseOptions>] [-AsResult] [<CommonParameters>]
```

## DESCRIPTION
Parses an AsciiDoc file or source string into OfficeIMO's native document model.

## EXAMPLES

### EXAMPLE 1
```powershell
Open-OfficeAsciiDoc -Path 'C:\Path'
```


### EXAMPLE 2
```powershell
Open-OfficeAsciiDoc -LiteralPath 'C:\Path'
```


### EXAMPLE 3
```powershell
Open-OfficeAsciiDoc -Text 'Value'
```


## PARAMETERS

### -AsResult
Return the parse result with diagnostics instead of only the document.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LiteralPath
Exact input path without wildcard expansion.

```yaml
Type: String
Parameter Sets: LiteralPath
Aliases: PSPath, FullName
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -Options
Optional parser settings.

```yaml
Type: AsciiDocParseOptions
Parameter Sets: Path, LiteralPath, Text
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
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
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

- `System.String`

## OUTPUTS

- `OfficeIMO.AsciiDoc.AsciiDocDocument`
- `OfficeIMO.AsciiDoc.AsciiDocParseResult`

## RELATED LINKS

- None
