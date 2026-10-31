---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Open-OfficeMarkdown
## SYNOPSIS
Opens native Markdown from a file or parses text; a file content block saves atomically.

## SYNTAX
### Path (Default)
```powershell
Open-OfficeMarkdown [-Path] <string> [[-Content] <scriptblock>] [-Options <Object[]>] [-WriteOptions <MarkdownSaveOptions>] [-PassThru] [-AllowLoss] [-Show] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPath
```powershell
Open-OfficeMarkdown [[-Content] <scriptblock>] -LiteralPath <string> [-Options <Object[]>] [-WriteOptions <MarkdownSaveOptions>] [-PassThru] [-AllowLoss] [-Show] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Text
```powershell
Open-OfficeMarkdown [[-Content] <scriptblock>] -Text <string> [-Options <Object[]>] [-PassThru] [-AllowLoss] [-Show] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Opens native Markdown from a file or parses text; a file content block saves atomically.

## EXAMPLES

### EXAMPLE 1
```powershell
$document = Get-Item -LiteralPath './Review[final].md' | Open-OfficeMarkdown
```


### EXAMPLE 2
```powershell
Open-OfficeMarkdown -Text '# Report' -Options (New-OfficeMarkdownReaderOptions -MaxInputCharacters 500000)
```


## PARAMETERS

### -AllowLoss
Gets or sets whether native-save loss is explicitly accepted.

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

### -Content
Gets or sets a content block; omission returns the live document.

```yaml
Type: ScriptBlock
Parameter Sets: Path, LiteralPath, Text
Aliases: None
Possible values:

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LiteralPath
Literal source path, including FileInfo.FullName and provider PSPath.

```yaml
Type: String
Parameter Sets: LiteralPath
Aliases: FullName, PSPath
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -Options
Native reader options followed by optional native-property hashtable overrides.

```yaml
Type: Object[]
Parameter Sets: Path, LiteralPath, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Gets or sets whether a completed content block returns the saved FileInfo.

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

### -Path
Markdown source with FileSystem wildcard expansion.

```yaml
Type: String
Parameter Sets: Path
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -Show
Gets or sets whether the saved file opens in the system default application.

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

### -Text
Markdown to parse without a file or authoring block.

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

### -WriteOptions
Native settings used when saving a file content block.

```yaml
Type: MarkdownSaveOptions
Parameter Sets: Path, LiteralPath
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

- `System.String`

## OUTPUTS

- `OfficeIMO.Markdown.MarkdownDocument`
- `System.IO.FileInfo`

## RELATED LINKS

- None
