---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Export-OfficeNotebook
## SYNOPSIS
Exports a live notebook or a selected saved structured output, returning the conversion report.

## SYNTAX
### Document (Default)
```powershell
Export-OfficeNotebook [-Path] <string> -Document <NotebookDocument> [-IncludeCode] [-IncludeOutputs] [-IncludeExecutionCounts] [-Title <string>] [-Author <string>] [-FailOnLoss] [-NoClobber] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### SavedTable
```powershell
Export-OfficeNotebook [-Path] <string> -Output <NotebookOutput> [-NoClobber] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Use .pdf, .html, .md, .docx or .ps1 for live notebooks. Pipe a saved structured table to write .csv or .xlsx. The command never runs cells and preserves an existing destination if conversion fails.

## EXAMPLES

### EXAMPLE 1
```powershell
$notebook | Export-OfficeNotebook -Path ./lab.pdf -IncludeCode:$false -Title 'Lab report'
```


### EXAMPLE 2
```powershell
$notebook | Get-OfficeNotebookCell -Index 2 | Get-OfficeNotebookOutput -Index 0 | Export-OfficeNotebook -Path ./results.xlsx
```


### EXAMPLE 3
```powershell
$notebook | Export-OfficeNotebook -Path ./lab.ps1 -FailOnLoss
```


## PARAMETERS

### -Author
Optional author for a rendered document.

```yaml
Type: String
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Live notebook to export to a document or a PowerShell script.

```yaml
Type: NotebookDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -FailOnLoss
Reject a notebook export that reports fidelity loss before publishing a file.

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

### -IncludeCode
Include code in a rendered document. Enabled by default. Use -IncludeCode:$false to omit it.

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

### -IncludeExecutionCounts
Include saved execution counts in a rendered document. Enabled by default.

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

### -IncludeOutputs
Include saved outputs in a rendered document. Enabled by default.

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

### -NoClobber
Refuse to replace an existing destination.

```yaml
Type: SwitchParameter
Parameter Sets: Document, SavedTable
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Output
Selected saved output containing an inline structured data resource.

```yaml
Type: NotebookOutput
Parameter Sets: SavedTable
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Path
Output path. The extension selects the format.

```yaml
Type: String
Parameter Sets: Document, SavedTable
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Title
Optional title for a rendered document.

```yaml
Type: String
Parameter Sets: Document
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
- `OfficeIMO.Notebook.NotebookOutput`

## OUTPUTS

- `OfficeIMO.OfficeConversionReport`

## RELATED LINKS

- None
