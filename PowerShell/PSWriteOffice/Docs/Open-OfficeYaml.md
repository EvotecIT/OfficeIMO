---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Open-OfficeYaml
## SYNOPSIS
Opens a YAML file, stream or text as a lossless live document; a file content block saves atomically.

## SYNTAX
### Path (Default)
```powershell
Open-OfficeYaml [-Path] <string> [[-Content] <scriptblock>] [-PassThru] [-Show] [-Options <YamlLoadOptions>] [-Untrusted] [-Strict] [-Schema <YamlSchema>] [-WriteOptions <YamlSaveOptions>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPath
```powershell
Open-OfficeYaml [[-Content] <scriptblock>] -LiteralPath <string> [-PassThru] [-Show] [-Options <YamlLoadOptions>] [-Untrusted] [-Strict] [-Schema <YamlSchema>] [-WriteOptions <YamlSaveOptions>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Text
```powershell
Open-OfficeYaml -Text <string> [-Options <YamlLoadOptions>] [-Untrusted] [-Strict] [-Schema <YamlSchema>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Stream
```powershell
Open-OfficeYaml -Stream <Stream> [-Options <YamlLoadOptions>] [-Untrusted] [-Strict] [-Schema <YamlSchema>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Saving writes only what changed: comments, blank lines, quoting and indentation outside the edited nodes stay
byte for byte. Malformed input fails with a diagnostic error ID (for example yaml-syntax-unexpected-token), and the
message gives the line, column and offending source line.

## EXAMPLES

### EXAMPLE 1
```powershell
Open-OfficeYaml -Path ./deployment.yaml {
    Set-OfficeYamlNode -Select 'spec.replicas' -Value 3
}
```


### EXAMPLE 2
```powershell
$document = Open-OfficeYaml -LiteralPath ./uploads/values.yaml -Untrusted -Strict
Get-OfficeYamlNode -Document $document -Select 'image.tag'
Close-OfficeYaml -Document $document
```


## PARAMETERS

### -Content
An optional editing block; omission returns the live document.

```yaml
Type: ScriptBlock
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LiteralPath
Literal YAML file, including FileInfo.FullName and provider PSPath.

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
Native load options or a hashtable of their properties: limits, schema, strictness and duplicate keys.

```yaml
Type: YamlLoadOptions
Parameter Sets: Path, LiteralPath, Text, Stream
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the FileInfo after a completed editing block.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
YAML file with FileSystem wildcard expansion.

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

### -Schema
Forces a schema for typed values: Failsafe, Json, Core (the default) or Yaml11.

```yaml
Type: YamlSchema
Parameter Sets: Path, LiteralPath, Text, Stream
Aliases: None
Possible values: Failsafe, Json, Core, Yaml11

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Show
Open the saved file after a completed editing block.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Stream
A readable stream, read from its current position; it stays open.

```yaml
Type: Stream
Parameter Sets: Stream
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Strict
Applies the YAML specification's exact rules instead of accepting common real-world leniencies.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath, Text, Stream
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Text
YAML text to parse without a file.

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

### -Untrusted
Caps every limit at the engine's untrusted defaults, for files from outside.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath, Text, Stream
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -WriteOptions
Native save settings used when an editing block saves; omission writes a minimal-change save.

```yaml
Type: YamlSaveOptions
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

- `OfficeIMO.Yaml.YamlDocument`
- `System.IO.FileInfo`

## RELATED LINKS

- None
