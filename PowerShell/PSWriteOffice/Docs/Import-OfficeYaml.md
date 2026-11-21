---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Import-OfficeYaml
## SYNOPSIS
Reads YAML files as PowerShell objects; the file form of ConvertFrom-OfficeYaml.

## SYNTAX
### Path (Default)
```powershell
Import-OfficeYaml [-Path] <string[]> [-AsHashtable] [-Depth <int>] [-NoEnumerate] [-AllDocuments] [-Schema <YamlSchema>] [-Strict] [-Untrusted] [<CommonParameters>]
```

### LiteralPath
```powershell
Import-OfficeYaml -LiteralPath <string[]> [-AsHashtable] [-Depth <int>] [-NoEnumerate] [-AllDocuments] [-Schema <YamlSchema>] [-Strict] [-Untrusted] [<CommonParameters>]
```

## DESCRIPTION
The encoding (UTF-8, UTF-16 or UTF-32) is detected from the file. Output follows ConvertFrom-OfficeYaml: mappings
become PSCustomObjects in file order, or ordered hashtables with AsHashtable. A file that fails to parse writes a
non-terminating error with the engine's code, line and column, and the next file is read.

## EXAMPLES

### EXAMPLE 1
```powershell
$settings = Import-OfficeYaml -Path ./settings.yaml
$settings.logging.level
```


### EXAMPLE 2
```powershell
Get-ChildItem ./deploy -Filter *.yaml | Import-OfficeYaml -AllDocuments | Group-Object kind
```


## PARAMETERS

### -AllDocuments
Read every document of a multi-document file; without it, more than one document is an error.

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

### -AsHashtable
Return mappings as ordered hashtables instead of PSCustomObjects.

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

### -Depth
Deepest nesting of mappings and sequences to accept; sets the reader's MaxDepth.

```yaml
Type: Int32
Parameter Sets: Path, LiteralPath
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

### -NoEnumerate
Write a root sequence as one array instead of enumerating its items.

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
Schema for scalar types: Core (the default), Json, Yaml11, or Failsafe to keep every scalar a string.

```yaml
Type: YamlSchema
Parameter Sets: Path, LiteralPath
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
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Untrusted
Caps every limit at the engine's untrusted defaults, for files from outside.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.String[]`

## OUTPUTS

- `System.Management.Automation.PSObject`
- `System.Collections.Specialized.OrderedDictionary`
- `System.Object[]`

## RELATED LINKS

- None
