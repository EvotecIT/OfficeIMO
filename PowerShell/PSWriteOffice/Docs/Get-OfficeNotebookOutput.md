---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficeNotebookOutput
## SYNOPSIS
Returns saved notebook outputs in cell order without running any code.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficeNotebookOutput -Cell <NotebookCell> [-Index <Int32>] [<CommonParameters>]
```

## DESCRIPTION
Returns saved notebook outputs in cell order without running any code.

## EXAMPLES

### EXAMPLE 1
```powershell
$notebook | Get-OfficeNotebookCell -Index 1 | Get-OfficeNotebookOutput -Index 0
```


## PARAMETERS

### -Cell
Cell containing the saved results.

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

### -Index
Zero-based saved output position. Omit to enumerate the available outputs.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Notebook.NotebookCell`

## OUTPUTS

- `OfficeIMO.Notebook.NotebookOutput`

## RELATED LINKS

- None
