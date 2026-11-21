---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# ConvertTo-OfficeYaml
## SYNOPSIS
Converts PowerShell objects to YAML text, as ConvertTo-Json does for JSON.

## SYNTAX
### InputObject (Default)
```powershell
ConvertTo-OfficeYaml [-InputObject] <Object> [-Depth <int>] [-WriteOptions <YamlWriterOptions>] [<CommonParameters>]
```

## DESCRIPTION
Hashtables, ordered dictionaries and object properties become mappings, collections become sequences and
enums their names. One piped object is written as itself and several as a sequence. Objects nested deeper than
Depth (default 2) are written as their string form with a warning, as ConvertTo-Json does. Strings that YAML 1.1
readers would take as another type (yes, on, 012, 1e3) are quoted.

## EXAMPLES

### EXAMPLE 1
```powershell
[ordered]@{ name = 'web'; replicas = 3; ports = 80, 443 } | ConvertTo-OfficeYaml
```


### EXAMPLE 2
```powershell
Get-Service | Select-Object -First 3 Name, Status, DependentServices |
                ConvertTo-OfficeYaml -Depth 3 -WriteOptions @{ MappingIndent = 4; SequenceIndent = 4; SequenceDashOffset = 2 }
```


## PARAMETERS

### -Depth
How many levels of nested objects to write; deeper values become their string form.

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
Objects to convert; an argument is written exactly as given, piped objects as one value or a sequence.

```yaml
Type: Object
Parameter Sets: InputObject
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -WriteOptions
Indentation and line-break settings, as a YamlWriterOptions object or hashtable.

```yaml
Type: YamlWriterOptions
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

- `System.Object`

## OUTPUTS

- `System.String`

## RELATED LINKS

- None
