---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Close-OfficeNotebook
## SYNOPSIS
Closes a notebook and releases its source-backed values without writing a file.

## SYNTAX
### __AllParameterSets
```powershell
Close-OfficeNotebook -Document <NotebookDocument> [<CommonParameters>]
```

## DESCRIPTION
Closes a notebook and releases its source-backed values without writing a file.

## EXAMPLES

### EXAMPLE 1
```powershell
$notebook | Close-OfficeNotebook
```


## PARAMETERS

### -Document
Live notebook to close.

```yaml
Type: NotebookDocument
Parameter Sets: __AllParameterSets
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

- `OfficeIMO.Notebook.NotebookDocument`

## OUTPUTS

- `None`

## RELATED LINKS

- None
