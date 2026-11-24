---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Open-OfficeNotebook
## SYNOPSIS
Opens a Jupyter notebook or imports a saved notebook source into the editable model.

## SYNTAX
### Path (Default)
```powershell
Open-OfficeNotebook [-Path] <string> [-Format <NotebookImportFormat>] [-CompanionPath <string>] [-DeepnoteNotebookId <string>] [-ScriptLanguage <string>] [-LoadOptions <NotebookLoadOptions>] [-FailOnLoss] [<CommonParameters>]
```

### LiteralPath
```powershell
Open-OfficeNotebook -LiteralPath <string> [-Format <NotebookImportFormat>] [-CompanionPath <string>] [-DeepnoteNotebookId <string>] [-ScriptLanguage <string>] [-LoadOptions <NotebookLoadOptions>] [-FailOnLoss] [<CommonParameters>]
```

## DESCRIPTION
Notebook cells remain inert. Close the returned document after using its source-backed values. Imported sources require an explicit .ipynb destination when saving.

## EXAMPLES

### EXAMPLE 1
```powershell
$notebook = Open-OfficeNotebook -LiteralPath ./lab.ipynb
$notebook | Get-OfficeNotebookCell -Kind Code
$notebook | Close-OfficeNotebook
```


### EXAMPLE 2
```powershell
Open-OfficeNotebook -LiteralPath ./lab.py -Format JupytextPercent -CompanionPath ./lab.ipynb -FailOnLoss
```


## PARAMETERS

### -CompanionPath
Literal path to a Jupyter notebook whose saved outputs may be matched to a Jupytext script.

```yaml
Type: String
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DeepnoteNotebookId
Exact notebook identifier within a Deepnote project.

```yaml
Type: String
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FailOnLoss
Reject source findings that describe fidelity loss.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Format
Explicit import profile for ambiguous source extensions. Omission uses native .ipynb or an unambiguous import extension.

```yaml
Type: NotebookImportFormat
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values: PolyglotDib, JupytextPercent, Quarto, MySTMarkdown, RMarkdown, RNotebookHtml, Deepnote, Zeppelin, PowerShellScript

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LiteralPath
Literal notebook path, including FileInfo.FullName or provider PSPath from a pipeline.

```yaml
Type: String
Parameter Sets: LiteralPath
Aliases: FullName, PSPath
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -LoadOptions
Native resource limits or a hashtable of limit properties, applied to input and companion files.

```yaml
Type: NotebookLoadOptions
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Notebook path with FileSystem wildcard expansion.

```yaml
Type: String
Parameter Sets: Path
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -ScriptLanguage
Optional language override for Jupytext percent comment markers.

```yaml
Type: String
Parameter Sets: Path, LiteralPath
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

- `System.String`

## OUTPUTS

- `OfficeIMO.Notebook.NotebookDocument`

## RELATED LINKS

- None
