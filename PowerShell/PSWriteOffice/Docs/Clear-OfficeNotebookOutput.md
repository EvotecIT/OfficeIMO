---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Clear-OfficeNotebookOutput
## SYNOPSIS
Clears saved results from a code cell, optionally including its execution count.

## SYNTAX
### __AllParameterSets
```powershell
Clear-OfficeNotebookOutput -Cell <NotebookCell> [-ExecutionCount] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Clears saved results from a code cell, optionally including its execution count.

## EXAMPLES

### EXAMPLE 1
```powershell
$notebook | Get-OfficeNotebookCell -Kind Code | Clear-OfficeNotebookOutput -ExecutionCount
$notebook | Save-OfficeNotebook -Path ./notebook-for-sharing.ipynb
```


## PARAMETERS

### -Cell
Code cell containing the saved results.

```yaml
Type: NotebookCell
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -ExecutionCount
Also clear the code cell's saved execution count.

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

### -PassThru
Return the updated cell.

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

- `OfficeIMO.Notebook.NotebookCell`

## OUTPUTS

- `OfficeIMO.Notebook.NotebookCell`

## RELATED LINKS

- None
