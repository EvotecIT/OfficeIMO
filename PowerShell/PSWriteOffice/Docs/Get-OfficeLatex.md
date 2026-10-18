---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeLatex
## SYNOPSIS
Parses a LaTeX file or source string into OfficeIMO's native document model.

## SYNTAX
### Path (Default)
```powershell
Get-OfficeLatex [-Path] <string> [-Options <LatexParseOptions>] [-AsResult] [<CommonParameters>]
```

### Text
```powershell
Get-OfficeLatex -Text <string> [-Options <LatexParseOptions>] [-AsResult] [<CommonParameters>]
```

## DESCRIPTION
Parses a LaTeX file or source string into OfficeIMO's native document model.

## EXAMPLES

### EXAMPLE 1
```powershell
Get-OfficeLatex -Path 'C:\Path'
```


### EXAMPLE 2
```powershell
Get-OfficeLatex -Text 'Value'
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
Type: LatexParseOptions
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
Path to a LaTeX file.

```yaml
Type: String
Parameter Sets: Path
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Text
LaTeX source text.

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

- `OfficeIMO.Latex.LatexDocument`
- `OfficeIMO.Latex.LatexParseResult`

## RELATED LINKS

- None
