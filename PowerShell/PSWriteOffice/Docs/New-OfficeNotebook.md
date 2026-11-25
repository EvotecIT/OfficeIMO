---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficeNotebook
## SYNOPSIS
Creates a live Jupyter notebook that can be populated and saved without running cells.

## SYNTAX
### __AllParameterSets
```powershell
New-OfficeNotebook [[-Path] <string>] [<CommonParameters>]
```

## DESCRIPTION
Creates a live Jupyter notebook that can be populated and saved without running cells.

## EXAMPLES

### EXAMPLE 1
```powershell
$notebook = New-OfficeNotebook -Path ./lab.ipynb
$notebook | Add-OfficeNotebookCell -Kind Markdown -Source '# Lab notes'
$notebook | Add-OfficeNotebookCell -Kind Code -Language PowerShell -Source 'Get-Date'
$notebook | Save-OfficeNotebook
$notebook | Close-OfficeNotebook
```


## PARAMETERS

### -Path
Associates a default .ipynb save destination without writing it.

```yaml
Type: String
Parameter Sets: __AllParameterSets
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

- `None`

## OUTPUTS

- `OfficeIMO.Notebook.NotebookDocument`

## RELATED LINKS

- None
