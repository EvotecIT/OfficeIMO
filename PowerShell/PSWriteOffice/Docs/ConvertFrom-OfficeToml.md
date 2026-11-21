---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# ConvertFrom-OfficeToml
## SYNOPSIS
Converts TOML text to PowerShell objects, as ConvertFrom-Json does for JSON.

## SYNTAX
### InputObject (Default)
```powershell
ConvertFrom-OfficeToml [-InputObject] <string> [-AsHashtable] [-Depth <int>] [-Untrusted] [<CommonParameters>]
```

## DESCRIPTION
Tables become PSCustomObjects with properties in file order (ordered hashtables with AsHashtable), arrays and
arrays of tables become arrays, integers Int64, floats Double, and date-times DateTimeOffset (with an offset),
DateTime (local date-time or date) or TimeSpan (local time). Piped lines are joined before parsing, so Get-Content
output works directly. Depth sets the deepest nesting accepted.

## EXAMPLES

### EXAMPLE 1
```powershell
$project = Get-Content -Path ./pyproject.toml -Raw | ConvertFrom-OfficeToml
$project.project.dependencies
```


### EXAMPLE 2
```powershell
(Get-Content ./Cargo.lock | ConvertFrom-OfficeToml -AsHashtable).package | ForEach-Object { '{0} {1}' -f $_.name, $_.version }
```


## PARAMETERS

### -AsHashtable
Return tables as ordered hashtables instead of PSCustomObjects.

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
Deepest nesting of tables and arrays to accept; sets the reader's MaxDepth.

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
TOML text; piped strings are joined with line breaks before parsing.

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

## RELATED LINKS

- None
