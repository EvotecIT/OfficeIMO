---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# ConvertTo-OfficeToml
## SYNOPSIS
Converts PowerShell objects to TOML text, as ConvertTo-Json does for JSON.

## SYNTAX
### InputObject (Default)
```powershell
ConvertTo-OfficeToml [-InputObject] <Object> [-Depth <int>] [-RootName <string>] [<CommonParameters>]
```

## DESCRIPTION
Hashtables, ordered dictionaries and object properties become tables: nested objects are written as
[table] sections and lists of objects as [[array]] sections, after the table's own values. Collections
become arrays, DateTime and DateTimeOffset TOML date-times, and enums their names. TOML has no null, so nulls are
omitted with a warning, and input that is not an object (such as several piped objects) is written under RootName.
Objects nested deeper than Depth (default 2) are written as their string form with a warning.

## EXAMPLES

### EXAMPLE 1
```powershell
[ordered]@{ name = 'web'; port = 8080; tls = [ordered]@{ enabled = $true } } | ConvertTo-OfficeToml
```


### EXAMPLE 2
```powershell
Get-Process | Select-Object -First 3 Name, Id | ConvertTo-OfficeToml -RootName process
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
Objects to convert; an argument is written exactly as given, piped objects as one value or an array.

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

### -RootName
The key that holds input which is not an object, because a TOML document is a table.

```yaml
Type: String
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
