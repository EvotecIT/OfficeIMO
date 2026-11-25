---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficeNotebookCell
## SYNOPSIS
Streams notebook cells in document order, optionally by position or cell kind.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficeNotebookCell -Document <NotebookDocument> [-Index <Int32>] [-Kind <NotebookCellKind>] [<CommonParameters>]
```

## DESCRIPTION
Streams notebook cells in document order, optionally by position or cell kind.

## EXAMPLES

### EXAMPLE 1
```powershell
$notebook | Get-OfficeNotebookCell -Kind Code | Get-OfficeNotebookOutput
```


## PARAMETERS

### -Document
Live notebook containing the cells.

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
Zero-based cell position. Omit to enumerate the notebook.

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

### -Kind
Optional cell-kind filter.

```yaml
Type: NotebookCellKind
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Unknown, Markdown, Code, Raw

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

- `OfficeIMO.Notebook.NotebookCell`

## RELATED LINKS

- None
