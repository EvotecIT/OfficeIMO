---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficeNotebookCell
## SYNOPSIS
Appends a Markdown, code, or raw cell to a live notebook.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficeNotebookCell [-Kind] <NotebookCellKind> [-Source] <string> -Document <NotebookDocument> [-Language <string>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Appends a Markdown, code, or raw cell to a live notebook.

## EXAMPLES

### EXAMPLE 1
```powershell
$cell = $notebook | Add-OfficeNotebookCell -Kind Code -Language PowerShell -Source 'Get-Date' -PassThru
```


## PARAMETERS

### -Document
Live notebook to extend.

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

### -Kind
Kind of cell to append. Choose Markdown, Code, or Raw.

```yaml
Type: NotebookCellKind
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Unknown, Markdown, Code, Raw

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Language
Explicit language for a code cell.

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

### -PassThru
Return the appended cell.

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
Cell text, which may be empty.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 1
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
