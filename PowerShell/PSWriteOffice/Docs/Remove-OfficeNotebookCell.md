---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Remove-OfficeNotebookCell
## SYNOPSIS
Removes a cell from a live notebook.

## SYNTAX
### __AllParameterSets
```powershell
Remove-OfficeNotebookCell [-Index] <int> -Document <NotebookDocument> [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Removes a cell from a live notebook.

## EXAMPLES

### EXAMPLE 1
```powershell
$notebook | Remove-OfficeNotebookCell -Index 0 -WhatIf
```


## PARAMETERS

### -Document
Live notebook containing the cell.

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

### -Index
Zero-based position of the cell to remove.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Return the updated document.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
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

- `OfficeIMO.Notebook.NotebookDocument`

## OUTPUTS

- `OfficeIMO.Notebook.NotebookDocument`

## RELATED LINKS

- None
