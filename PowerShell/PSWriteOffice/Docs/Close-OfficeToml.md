---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Close-OfficeToml
## SYNOPSIS
Completes a live TOML document's lifetime without saving.

## SYNTAX
### Document (Default)
```powershell
Close-OfficeToml -Document <TomlDocument> [<CommonParameters>]
```

## DESCRIPTION
Completes a live TOML document's lifetime without saving.

## EXAMPLES

### EXAMPLE 1
```powershell
Close-OfficeToml -Document $document
```


### EXAMPLE 2
```powershell
$document = Open-OfficeToml -Path ./Cargo.toml
(Get-OfficeTomlNode -Document $document -Select 'package.version').Value
Close-OfficeToml -Document $document
```


## PARAMETERS

### -Document
TOML document to close.

```yaml
Type: TomlDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
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
