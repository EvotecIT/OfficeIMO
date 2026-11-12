---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# ConvertFrom-OfficeYaml
## SYNOPSIS
Converts YAML text to PowerShell objects, as ConvertFrom-Json does for JSON.

## SYNTAX
### InputObject (Default)
```powershell
ConvertFrom-OfficeYaml [-InputObject] <string> [-AsHashtable] [-Depth <int>] [-NoEnumerate] [-AllDocuments] [-Schema <YamlSchema>] [-Strict] [-Untrusted] [<CommonParameters>]
```

## DESCRIPTION
Mappings become PSCustomObjects with properties in file order (ordered hashtables with AsHashtable), sequences
become arrays, and scalars resolve through the Core schema to Int64, BigInteger, Double, Boolean, $null or String.
Schema Failsafe keeps every scalar a string; Schema Yaml11 reads yes/no/on/off as Booleans. Aliases and merge keys
are resolved. Piped lines are joined before parsing, so Get-Content output works directly. Depth sets the deepest
nesting accepted (the engine default is 256).

## EXAMPLES

### EXAMPLE 1
```powershell
$config = Get-Content -Path ./settings.yaml -Raw | ConvertFrom-OfficeYaml
$config.database.port
```


### EXAMPLE 2
```powershell
Get-Content ./manifests.yaml | ConvertFrom-OfficeYaml -AllDocuments -AsHashtable |
                ForEach-Object { '{0}/{1}' -f $_.kind, $_.metadata.name }
```


## PARAMETERS

### -AllDocuments
Read every document of a multi-document stream; without it, more than one document is an error.

```yaml
Type: SwitchParameter
Parameter Sets: InputObject
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
Parameter Sets: InputObject
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
Parameter Sets: InputObject
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -InputObject
YAML text; piped strings are joined with line breaks before parsing.

```yaml
Type: String
Parameter Sets: InputObject
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -NoEnumerate
Write a root sequence as one array instead of enumerating its items.

```yaml
Type: SwitchParameter
Parameter Sets: InputObject
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Schema
Schema for scalar types: Core (the default), Json, Yaml11, or Failsafe to keep every scalar a string.

```yaml
Type: YamlSchema
Parameter Sets: InputObject
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
Parameter Sets: InputObject
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Untrusted
Caps every limit at the engine's untrusted defaults, for text from outside.

```yaml
Type: SwitchParameter
Parameter Sets: InputObject
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

- `System.Management.Automation.PSObject`
- `System.Collections.Specialized.OrderedDictionary`
- `System.Object[]`

## RELATED LINKS

- None
