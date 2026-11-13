---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Export-OfficeToml
## SYNOPSIS
Writes PowerShell objects to a TOML file; the file form of ConvertTo-OfficeToml.

## SYNTAX
### InputObject (Default)
```powershell
Export-OfficeToml [-Path] <string> -InputObject <Object> [-Depth <int>] [-RootName <string>] [-NoClobber] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
The file is UTF-8 without a byte-order mark and is written atomically. An existing file is replaced unless
NoClobber is set. Nested objects become [table] sections and lists of objects [[array]] sections; nulls
are omitted with a warning.

## EXAMPLES

### EXAMPLE 1
```powershell
[ordered]@{ title = 'Report'; owner = [ordered]@{ name = 'Ops'; updated = Get-Date } } | Export-OfficeToml -Path ./settings.toml
```


### EXAMPLE 2
```powershell
Import-OfficeCsv -Path ./servers.csv | Export-OfficeToml -Path ./inventory.toml -RootName server -NoClobber -PassThru
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
Objects to write; an argument is written exactly as given, piped objects as one value or an array.

```yaml
Type: Object
Parameter Sets: InputObject
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -NoClobber
Refuse to replace an existing file.

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

### -PassThru
Return the written file.

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

### -Path
The TOML file to create.

```yaml
Type: String
Parameter Sets: InputObject
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
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

- `System.IO.FileInfo`

## RELATED LINKS

- None
