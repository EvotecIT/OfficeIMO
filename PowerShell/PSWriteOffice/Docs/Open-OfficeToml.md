---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Open-OfficeToml
## SYNOPSIS
Opens a TOML file, stream or text as a lossless live document; a file content block saves atomically.

## SYNTAX
### Path (Default)
```powershell
Open-OfficeToml [-Path] <string> [[-Content] <scriptblock>] [-PassThru] [-Show] [-Options <TomlLoadOptions>] [-Untrusted] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPath
```powershell
Open-OfficeToml [[-Content] <scriptblock>] -LiteralPath <string> [-PassThru] [-Show] [-Options <TomlLoadOptions>] [-Untrusted] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Text
```powershell
Open-OfficeToml -Text <string> [-Options <TomlLoadOptions>] [-Untrusted] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Stream
```powershell
Open-OfficeToml -Stream <Stream> [-Options <TomlLoadOptions>] [-Untrusted] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Saving writes only what changed: comments, blank lines, quoting, key order, alignment and the spelling of
every other value stay byte for byte. Malformed input fails with an error ID built from the engine's code (for
example toml-Toml.Load.DuplicateKey), and the message gives the line and column.

## EXAMPLES

### EXAMPLE 1
```powershell
Open-OfficeToml -Path ./pyproject.toml {
    Set-OfficeTomlNode -Select 'project.version' -Value '1.2.0'
}
```


### EXAMPLE 2
```powershell
$document = Open-OfficeToml -LiteralPath ./uploads/Cargo.toml -Untrusted
Get-OfficeTomlNode -Document $document -Select 'package.name'
Close-OfficeToml -Document $document
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
Literal TOML file, including FileInfo.FullName and provider PSPath.

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
Native load options or a hashtable of their properties: MaxDepth, MaxNodes, MaxStringLength and MaxInputBytes.

```yaml
Type: TomlLoadOptions
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
TOML file with FileSystem wildcard expansion.

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

### -Text
TOML text to parse without a file.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.String`

## OUTPUTS

- `OfficeIMO.Toml.TomlDocument`
- `System.IO.FileInfo`

## RELATED LINKS

- None
