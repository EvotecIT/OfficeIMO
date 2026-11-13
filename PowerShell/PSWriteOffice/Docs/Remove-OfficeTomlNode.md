---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Remove-OfficeTomlNode
## SYNOPSIS
Removes keys, tables or array items from a live TOML document by path.

## SYNTAX
### Document (Default)
```powershell
Remove-OfficeTomlNode [-Select] <string[]> [-Document <TomlDocument>] [<CommonParameters>]
```

## DESCRIPTION
A key goes with its line and the comment lines directly above it; a table goes with all of its sections,
wherever they are in the file. Neighboring keys and comments stay. A path with nothing at it writes a
non-terminating ObjectNotFound error.

## EXAMPLES

### EXAMPLE 1
```powershell
Open-OfficeToml -Path ./pyproject.toml {
    Remove-OfficeTomlNode -Select 'tool.ruff.lint.ignore'
}
```


### EXAMPLE 2
```powershell
Remove-OfficeTomlNode -Document $document -Select 'bin[1]' -ErrorAction SilentlyContinue
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

### -Select
Paths of the keys, tables or items to remove.

```yaml
Type: String[]
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
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

- `None`

## RELATED LINKS

- None
