---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Set-OfficeNotebookCell
## SYNOPSIS
Updates cell text or a saved code-cell execution count without running code.

## SYNTAX
### __AllParameterSets
```powershell
Set-OfficeNotebookCell -Cell <NotebookCell> [-Source <string>] [-ExecutionCount <Int32>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Updates cell text or a saved code-cell execution count without running code.

## EXAMPLES

### EXAMPLE 1
```powershell
$notebook | Get-OfficeNotebookCell -Index 1 | Set-OfficeNotebookCell -Source 'Get-Date' -ExecutionCount $null
```


## PARAMETERS

### -Cell
Cell to edit. Keep its notebook open until the edit is complete.

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
Saved execution count for a code cell. Supply $null to clear it.

```yaml
Type: Int32
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

### -Source
Replacement cell text.

```yaml
Type: String
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
