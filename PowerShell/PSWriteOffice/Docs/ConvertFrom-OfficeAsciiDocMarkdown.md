---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# ConvertFrom-OfficeAsciiDocMarkdown
## SYNOPSIS
Converts Markdown to native AsciiDoc with fidelity diagnostics.

## SYNTAX
### Path (Default)
```powershell
ConvertFrom-OfficeAsciiDocMarkdown [-Path] <string> [-OutputPath <string>] [-Options <MarkdownToAsciiDocOptions>] [-FailOnLoss] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
ConvertFrom-OfficeAsciiDocMarkdown -Document <MarkdownDoc> [-OutputPath <string>] [-Options <MarkdownToAsciiDocOptions>] [-FailOnLoss] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Converts Markdown to native AsciiDoc with fidelity diagnostics.

## EXAMPLES

### EXAMPLE 1
```powershell
ConvertFrom-OfficeAsciiDocMarkdown -Path 'C:\Path'
```


### EXAMPLE 2
```powershell
ConvertFrom-OfficeAsciiDocMarkdown -Document 'Value'
```


## PARAMETERS

### -Document
Markdown document to convert.

```yaml
Type: MarkdownDoc
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -FailOnLoss
Throw when a source feature cannot be mapped exactly.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Options
Optional conversion settings.

```yaml
Type: MarkdownToAsciiDocOptions
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputPath
Optional AsciiDoc destination path.

```yaml
Type: String
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Path to a Markdown file.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Markdown.MarkdownDoc`

## OUTPUTS

- `OfficeIMO.AsciiDoc.Markdown.MarkdownToAsciiDocResult`

## RELATED LINKS

- None
