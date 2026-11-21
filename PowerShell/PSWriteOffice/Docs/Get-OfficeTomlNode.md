---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficeTomlNode
## SYNOPSIS
Gets nodes of a live TOML document by path, as Path, Value, Node, Line and Column objects.

## SYNTAX
### Document (Default)
```powershell
Get-OfficeTomlNode [[-Select] <string[]>] [-AsHashtable] [-Document <TomlDocument>] [<CommonParameters>]
```

## DESCRIPTION
Select uses the engine's path grammar, the one YAML node paths use: dotted keys, [n] for 0-based array items
and tables of an array of tables, and ["key.with.dots"] for keys that are not plain names. Value holds plain
PowerShell data as ConvertFrom-OfficeToml returns it; Node is the live node. Line and Column are 1-based and point at
the value or the [ of a header; tables made by dotted keys and nodes added by editing report 0. Without Select, the
root table is returned.

## EXAMPLES

### EXAMPLE 1
```powershell
(Get-OfficeTomlNode -Document $document -Select 'tool.ruff.line-length').Value
```


### EXAMPLE 2
```powershell
Get-OfficeTomlNode -Document $document -Select 'project.version', 'tool.uv.sources["my.pkg"]' |
                Select-Object -Property Path, Value, Line, Column
```


## PARAMETERS

### -AsHashtable
Return tables in Value as ordered hashtables instead of PSCustomObjects.

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
Type: TomlDocument
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
Paths of the nodes to get; omission returns the root table.

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

- `OfficeIMO.Toml.TomlDocument`

## OUTPUTS

- `System.Management.Automation.PSObject`

## RELATED LINKS

- None
