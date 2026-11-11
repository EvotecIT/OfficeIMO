---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficeYamlNode
## SYNOPSIS
Gets nodes of a live YAML document by path, as Path, Value, Node, Line and Column objects.

## SYNTAX
### Document (Default)
```powershell
Get-OfficeYamlNode [[-Select] <string[]>] [-AsHashtable] [-Document <YamlDocument>] [<CommonParameters>]
```

## DESCRIPTION
Select uses the engine's path grammar: dotted keys, [n] for 0-based sequence items, ["key.with.dots"] for keys
that are not plain names, and a leading $[n] for the 0-based document in a multi-document stream. Value holds plain
PowerShell data as ConvertFrom-OfficeYaml returns it; Node is the live node. Line and Column are 1-based. Without
Select, each document's root is returned.

## EXAMPLES

### EXAMPLE 1
```powershell
(Get-OfficeYamlNode -Document $document -Select 'spec.template.spec.containers[0].image').Value
```


### EXAMPLE 2
```powershell
Get-OfficeYamlNode -Document $document -Select 'spec.replicas', 'metadata.labels["app.kubernetes.io/name"]' |
                Select-Object -Property Path, Value, Line, Column
```


## PARAMETERS

### -AsHashtable
Return mappings in Value as ordered hashtables instead of PSCustomObjects.

```yaml
Type: SwitchParameter
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Gets or sets an explicit live document; omission uses the enclosing content block.

```yaml
Type: YamlDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Select
Paths of the nodes to get; omission returns each document's root.

```yaml
Type: String[]
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Yaml.YamlDocument`

## OUTPUTS

- `System.Management.Automation.PSObject`

## RELATED LINKS

- None
