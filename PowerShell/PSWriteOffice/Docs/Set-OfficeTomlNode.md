---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Set-OfficeTomlNode
## SYNOPSIS
Sets the value at a path in a live TOML document, adding missing tables on the way.

## SYNTAX
### Document (Default)
```powershell
Set-OfficeTomlNode [-Select] <string> [-Value] <Object> [-PassThru] [-Document <TomlDocument>] [<CommonParameters>]
```

## DESCRIPTION
Setting a value over a value changes only its text, and an equal value keeps its original spelling, so saving
changes only what really changed. A new key is written after its table's last key in the indentation and spacing of
its neighbors; a missing table gets a [header] section. Hashtables and objects become inline tables, and arrays
become arrays. TOML has no null, so $null is an error and nulls inside a value are omitted with a warning. An index
equal to an array's length appends.

## EXAMPLES

### EXAMPLE 1
```powershell
Open-OfficeToml -Path ./Cargo.toml {
    Set-OfficeTomlNode -Select 'package.version' -Value '0.4.0'
}
```


### EXAMPLE 2
```powershell
Set-OfficeTomlNode -Document $document -Select 'dependencies.serde' -Value ([ordered]@{ version = '1'; features = @('derive') })
            Set-OfficeTomlNode -Document $document -Select 'project.dependencies[0]' -Value 'httpx>=0.27'
```


## PARAMETERS

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

### -PassThru
Return the node that is now at the path.

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

### -Select
Path of the key or array item to set.

```yaml
Type: String
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Value
The new value: a string, number, Boolean, date, hashtable, array or object.

```yaml
Type: Object
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: 1
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
