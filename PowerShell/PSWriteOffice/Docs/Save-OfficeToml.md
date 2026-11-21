---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Save-OfficeToml
## SYNOPSIS
Saves a live TOML document and keeps it available.

## SYNTAX
### Document (Default)
```powershell
Save-OfficeToml [[-Path] <string>] -Document <TomlDocument> [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
The save is minimal-change: untouched text is copied from the source, with its line endings and UTF-8
byte-order mark, and only edited keys, values and comments are written anew in the layout of their neighbors.

## EXAMPLES

### EXAMPLE 1
```powershell
Save-OfficeToml -Document $document
```


### EXAMPLE 2
```powershell
Save-OfficeToml -Document $document -Path ./out/pyproject.toml -PassThru
```


## PARAMETERS

### -Document
TOML document to save.

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

### -PassThru
Return the saved file.

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

### -Path
Optional destination, which becomes the document's FilePath; omission saves to its FilePath.

```yaml
Type: String
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

- `System.IO.FileInfo`

## RELATED LINKS

- None
