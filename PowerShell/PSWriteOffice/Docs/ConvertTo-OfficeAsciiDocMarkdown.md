---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# ConvertTo-OfficeAsciiDocMarkdown
## SYNOPSIS
Converts AsciiDoc to Markdown with fidelity diagnostics.

## SYNTAX
### Path (Default)
```powershell
ConvertTo-OfficeAsciiDocMarkdown [-Path] <string> [-OutputPath <string>] [-Options <AsciiDocToMarkdownOptions>] [-FailOnLoss] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
ConvertTo-OfficeAsciiDocMarkdown -Document <AsciiDocDocument> [-OutputPath <string>] [-Options <AsciiDocToMarkdownOptions>] [-FailOnLoss] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Converts AsciiDoc to Markdown with fidelity diagnostics.

## EXAMPLES

### EXAMPLE 1
```powershell
ConvertTo-OfficeAsciiDocMarkdown -Path 'C:\Path'
```


### EXAMPLE 2
```powershell
ConvertTo-OfficeAsciiDocMarkdown -Document 'Value'
```


## PARAMETERS

### -Document
AsciiDoc document to convert.

```yaml
Type: AsciiDocDocument
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
Type: AsciiDocToMarkdownOptions
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
Optional Markdown destination path.

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
Path to an AsciiDoc file.

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

- `OfficeIMO.AsciiDoc.AsciiDocDocument`

## OUTPUTS

- `OfficeIMO.AsciiDoc.Markdown.AsciiDocToMarkdownResult`

## RELATED LINKS

- None
