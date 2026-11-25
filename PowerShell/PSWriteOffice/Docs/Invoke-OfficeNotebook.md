---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Invoke-OfficeNotebook
## SYNOPSIS
Runs selected PowerShell notebook cells in one persistent session and records their outputs.

## SYNTAX
### Path (Default)
```powershell
Invoke-OfficeNotebook [-Path] <string> [-CellIndex <int[]>] [-TimeoutSeconds <int>] [-ContinueOnError] [-IncludeUnspecifiedLanguage] [-Isolation <PowerShellNotebookIsolation>] [-PowerShellPath <string>] [-WorkingDirectory <string>] [-StreamOutput] [-OutputPath <string>] [-NoClobber] [-AllowLoss] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPath
```powershell
Invoke-OfficeNotebook -LiteralPath <string> [-CellIndex <int[]>] [-TimeoutSeconds <int>] [-ContinueOnError] [-IncludeUnspecifiedLanguage] [-Isolation <PowerShellNotebookIsolation>] [-PowerShellPath <string>] [-WorkingDirectory <string>] [-StreamOutput] [-OutputPath <string>] [-NoClobber] [-AllowLoss] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
Invoke-OfficeNotebook -Document <NotebookDocument> [-CellIndex <int[]>] [-TimeoutSeconds <int>] [-ContinueOnError] [-IncludeUnspecifiedLanguage] [-Isolation <PowerShellNotebookIsolation>] [-PowerShellPath <string>] [-WorkingDirectory <string>] [-StreamOutput] [-OutputPath <string>] [-NoClobber] [-AllowLoss] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Run notebooks whose code you trust. Cells run under your account in a new runspace of the current PowerShell process, so they share its environment variables and loaded assemblies but not its variables or functions. Use -Isolation Process or -PowerShellPath to run them in a separate PowerShell process instead. The returned report contains the updated live notebook. Close the notebook when you finish using it.

Use -WhatIf to preview execution and any requested save. Use -Confirm to ask before running the notebook.

## EXAMPLES

### EXAMPLE 1
```powershell
$report = Invoke-OfficeNotebook -LiteralPath ./lab.ipynb -OutputPath ./executed.ipynb -NoClobber
if ($null -ne $report) {
    try {
        $report.Cells
    } finally {
        $report.Notebook | Close-OfficeNotebook
    }
}
```


### EXAMPLE 2
```powershell
Invoke-OfficeNotebook -LiteralPath ./lab.ipynb -OutputPath ./reviewed.ipynb -WhatIf
$report = Invoke-OfficeNotebook -LiteralPath ./lab.ipynb -OutputPath ./reviewed.ipynb -Confirm
if ($null -ne $report) {
    $report.Notebook | Close-OfficeNotebook
}
```


### EXAMPLE 3
```powershell
$notebook | Invoke-OfficeNotebook -CellIndex 0, 2 -TimeoutSeconds 15
```


### EXAMPLE 4
```powershell
$report = Invoke-OfficeNotebook -LiteralPath ./lab.ipynb -PowerShellPath "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
if ($null -ne $report) {
    $report.Notebook | Close-OfficeNotebook
}
```


## PARAMETERS

### -AllowLoss
Accept loss reported when normalizing imported content for the .ipynb output.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -CellIndex
Zero-based cell indexes. Selected PowerShell cells run in notebook order.

```yaml
Type: Int32[]
Parameter Sets: Path, LiteralPath, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ContinueOnError
Continue after PowerShell cell errors. A timeout or lost session stops execution.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Live notebook to update with the execution results.

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

### -IncludeUnspecifiedLanguage
Run cells without a declared cell or notebook language as PowerShell.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Isolation
Where cells run. Runspace, the default, uses a new runspace in this PowerShell process. Process starts a separate PowerShell, which a timeout or Ctrl+C ends with any programs it started. Choose Process for cells that change process-wide state or can block in calls PowerShell cannot stop.

```yaml
Type: PowerShellNotebookIsolation
Parameter Sets: Path, LiteralPath, Document
Aliases: None
Possible values: Process, Runspace

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LiteralPath
Literal .ipynb path, including a FileInfo.FullName or provider PSPath from the pipeline.

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

### -NoClobber
Refuse to replace an existing output file.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputPath
Atomically save the executed notebook to this .ipynb path, including any partial results reported after a cell error or timeout. Use this parameter with one notebook per invocation.

```yaml
Type: String
Parameter Sets: Path, LiteralPath, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Path to one native .ipynb notebook. FileSystem wildcards must resolve to a single file.

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

### -PowerShellPath
Literal path to the PowerShell executable for a separate process. Implies -Isolation Process. The default uses the installed executable for this host when available.

```yaml
Type: String
Parameter Sets: Path, LiteralPath, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -StreamOutput
Write live output records to the pipeline before the final execution report.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TimeoutSeconds
Maximum seconds per cell. The default is 60. Set to 0 to run without a per-cell timeout.

```yaml
Type: Int32
Parameter Sets: Path, LiteralPath, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -WorkingDirectory
Initial working directory. The default uses the notebook's directory when available.

```yaml
Type: String
Parameter Sets: Path, LiteralPath, Document
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
- `OfficeIMO.Notebook.NotebookDocument`

## OUTPUTS

- `OfficeIMO.Notebook.PowerShell.NotebookExecutionReport`
- `OfficeIMO.Notebook.PowerShell.NotebookExecutionEvent`

## RELATED LINKS

- None
