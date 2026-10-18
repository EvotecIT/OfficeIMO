---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeExcelPreflight
## SYNOPSIS
Runs OfficeIMO Excel feature and workflow preflight checks.

## SYNTAX
### Path (Default)
```powershell
Get-OfficeExcelPreflight [-Path] <string> [-Capability <ExcelPreflightCapability[]>] [-IncludeFeatures] [-IncludeRepairHints] [-AsMarkdown] [-ThrowOnFailure] [<CommonParameters>]
```

### Document
```powershell
Get-OfficeExcelPreflight -Document <ExcelDocument> [-Capability <ExcelPreflightCapability[]>] [-IncludeFeatures] [-IncludeRepairHints] [-AsMarkdown] [-ThrowOnFailure] [<CommonParameters>]
```

## DESCRIPTION
Runs OfficeIMO Excel feature and workflow preflight checks.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $preflight = Get-OfficeExcelPreflight -Path .\Report.xlsx -Capability EditCellValues,ExportPdfReport -IncludeFeatures -IncludeRepairHints
$preflight.Capabilities |
    Where-Object Passed -eq $false
```

Returns reusable OfficeIMO capability diagnostics and discovered workbook features.

## PARAMETERS

### -AsMarkdown
Return OfficeIMO's Markdown preflight report instead of structured objects.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Capability
Capabilities to evaluate or enforce. Defaults to all capabilities.

```yaml
Type: ExcelPreflightCapability[]
Parameter Sets: Path, Document
Aliases: None
Possible values: ReadWorkbookData, EditCellValues, EditWorkbookStructure, UseCachedFormulaValues, CalculateFormulas, BindTemplate, ExportPdfReport

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Workbook to inspect.

```yaml
Type: ExcelDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -IncludeFeatures
Include discovered feature rows in object output.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IncludeRepairHints
Include actionable OfficeIMO repair hints for blocked capabilities.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Path to the workbook.

```yaml
Type: String
Parameter Sets: Path
Aliases: InputPath, FilePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ThrowOnFailure
Throw when any requested capability is unavailable. Without -Capability, throws when advanced features need review.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Document
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

- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `System.Management.Automation.PSObject`
- `System.String`

## RELATED LINKS

- None
