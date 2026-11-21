---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Test-OfficeYaml
## SYNOPSIS
Tests whether YAML files or text load, without changing them.

## SYNTAX
### Path (Default)
```powershell
Test-OfficeYaml [-Path] <string[]> [-Options <YamlLoadOptions>] [-Untrusted] [-Strict] [-Schema <YamlSchema>] [-Detailed] [<CommonParameters>]
```

### LiteralPath
```powershell
Test-OfficeYaml -LiteralPath <string[]> [-Options <YamlLoadOptions>] [-Untrusted] [-Strict] [-Schema <YamlSchema>] [-Detailed] [<CommonParameters>]
```

### Text
```powershell
Test-OfficeYaml -Text <string> [-Options <YamlLoadOptions>] [-Untrusted] [-Strict] [-Schema <YamlSchema>] [-Detailed] [<CommonParameters>]
```

## DESCRIPTION
Returns $true or $false for each input, as Test-Json does. Detailed returns a result with the error's
diagnostic code, message, 1-based line and column, and the warnings the engine recorded while loading. Strict applies the
specification's exact rules, and Untrusted the engine's limits for files from outside.

## EXAMPLES

### EXAMPLE 1
```powershell
if (-not (Test-OfficeYaml -Path ./values.yaml -Strict)) { throw 'values.yaml is not valid YAML.' }
```


### EXAMPLE 2
```powershell
Get-ChildItem ./deploy -Recurse -Filter *.yaml | Test-OfficeYaml -Detailed |
                Where-Object { -not $_.IsValid } |
                Select-Object -Property Path, Line, Column, Code, Message
```


## PARAMETERS

### -Detailed
Return a result with the error position and the load warnings instead of a Boolean.

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
Literal YAML files, including FileInfo.FullName and provider PSPath.

```yaml
Type: String[]
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
Native load options or a hashtable of their properties.

```yaml
Type: YamlLoadOptions
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
YAML files with FileSystem wildcard expansion.

```yaml
Type: String[]
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
Forces a schema: Failsafe, Json, Core (the default) or Yaml11.

```yaml
Type: YamlSchema
Parameter Sets: Path, LiteralPath, Text
Aliases: None
Possible values: Failsafe, Json, Core, Yaml11

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Strict
Applies the YAML specification's exact rules instead of accepting common real-world leniencies.

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
YAML text to test.

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
Caps every limit at the engine's untrusted defaults.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.String[]`

## OUTPUTS

- `System.Boolean`
- `PSWriteOffice.Models.Yaml.OfficeYamlTestResult`: The outcome of Test-OfficeYaml -Detailed for one input.

## RELATED LINKS

- None
