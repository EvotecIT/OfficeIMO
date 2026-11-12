---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Export-OfficeYaml
## SYNOPSIS
Writes PowerShell objects to a YAML file; the file form of ConvertTo-OfficeYaml.

## SYNTAX
### InputObject (Default)
```powershell
Export-OfficeYaml [-Path] <string> -InputObject <Object> [-Depth <int>] [-WriteOptions <YamlWriterOptions>] [-NoClobber] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
The file is UTF-8 without a byte-order mark and is written atomically. An existing file is replaced unless
NoClobber is set. To change values in an existing file while keeping its comments and layout, use Open-OfficeYaml
and Set-OfficeYamlNode instead.

## EXAMPLES

### EXAMPLE 1
```powershell
[ordered]@{ logging = [ordered]@{ level = 'Information' }; retries = 3 } | Export-OfficeYaml -Path ./settings.yaml
```


### EXAMPLE 2
```powershell
Import-OfficeCsv -Path ./servers.csv | Export-OfficeYaml -Path ./inventory.yaml -NoClobber -PassThru
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
Objects to write; an argument is written exactly as given, piped objects as one value or a sequence.

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
The YAML file to create.

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

- `System.IO.FileInfo`

## RELATED LINKS

- None
