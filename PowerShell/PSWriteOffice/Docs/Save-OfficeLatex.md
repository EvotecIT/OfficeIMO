---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Save-OfficeLatex
## SYNOPSIS
Saves an OfficeIMO LaTeX document.

## SYNTAX
### __AllParameterSets
```powershell
Save-OfficeLatex [-Document] <LatexDocument> [-Path] <string> [-Options <LatexWriterOptions>] [-Mode <LatexWriterMode>] [-LineEnding <OfficeLineEnding>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Saves an OfficeIMO LaTeX document.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $document = Get-OfficeLatex -Path .\Article.tex
$document | Save-OfficeLatex -Path .\Article-normalized.tex -Mode Canonical
```


## PARAMETERS

### -Document
LaTeX document to save.

```yaml
Type: LatexDocument
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
Writer mode. Preserve retains unchanged source; Canonical normalizes output.

```yaml
Type: LatexWriterMode
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
Type: LatexWriterOptions
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

- `OfficeIMO.Latex.LatexDocument`

## OUTPUTS

- `OfficeIMO.Latex.LatexDocument`

## RELATED LINKS

- None
