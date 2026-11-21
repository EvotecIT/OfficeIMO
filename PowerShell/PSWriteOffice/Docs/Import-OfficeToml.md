---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Import-OfficeToml
## SYNOPSIS
Reads TOML files as PowerShell objects; the file form of ConvertFrom-OfficeToml.

## SYNTAX
### Path (Default)
```powershell
Import-OfficeToml [-Path] <string[]> [-AsHashtable] [-Depth <int>] [-Untrusted] [<CommonParameters>]
```

### LiteralPath
```powershell
Import-OfficeToml -LiteralPath <string[]> [-AsHashtable] [-Depth <int>] [-Untrusted] [<CommonParameters>]
```

## DESCRIPTION
Output follows ConvertFrom-OfficeToml: tables become PSCustomObjects in file order, or ordered hashtables with
AsHashtable. A file that fails to parse writes a non-terminating error with the engine's code, line and column, and
the next file is read.

## EXAMPLES

### EXAMPLE 1
```powershell
$ruff = (Import-OfficeToml -Path ./pyproject.toml).tool.ruff
$ruff.'line-length'
```


### EXAMPLE 2
```powershell
Get-ChildItem ./crates -Filter Cargo.toml -Recurse | Import-OfficeToml | ForEach-Object { $_.package.name }
```


## PARAMETERS

### -AsHashtable
Return tables as ordered hashtables instead of PSCustomObjects.

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
Deepest nesting of tables and arrays to accept; sets the reader's MaxDepth.

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
Literal TOML files, including FileInfo.FullName and provider PSPath.

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

### -Path
TOML files with FileSystem wildcard expansion.

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

## RELATED LINKS

- None
