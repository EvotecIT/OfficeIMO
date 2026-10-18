---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Join-OfficeExcelWorkbook
## SYNOPSIS
Merges worksheets from one or more workbooks into a target workbook.

## SYNTAX
### Path (Default)
```powershell
Join-OfficeExcelWorkbook [-Path] <string> [[-SourcePath] <string[]>] [-SourceDocument <ExcelDocument>] [-SourceSheet <string[]>] [-SheetNamePrefix <string>] [-ValidationMode <ExcelSheetNameValidationMode>] [-CopyMode <ExcelWorksheetCopyMode>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
Join-OfficeExcelWorkbook [[-SourcePath] <string[]>] -Document <ExcelDocument> [-SourceDocument <ExcelDocument>] [-SourceSheet <string[]>] [-SheetNamePrefix <string>] [-ValidationMode <ExcelSheetNameValidationMode>] [-CopyMode <ExcelWorksheetCopyMode>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Merges worksheets from one or more workbooks into a target workbook.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $sources = Get-ChildItem .\Incoming\*.xlsx | Select-Object -ExpandProperty FullName
$results = Join-OfficeExcelWorkbook -Path .\Combined.xlsx -SourcePath $sources -CopyMode Package -SheetNamePrefix Import
$results | Select-Object SheetCount, SourceSheets, TargetSheets
```

Copies worksheets between packages without importing rows into PowerShell objects, which is the preferred path for large workbook merge workflows.

### EXAMPLE 2
```powershell
PS> $merge = Join-OfficeExcelWorkbook -Path .\Target.xlsx -SourcePath .\Source.xlsx -SourceSheet Data,Summary -SheetNamePrefix 'Imported '
Get-OfficeExcelSummary -Path .\Target.xlsx |
    Select-Object Path, WorksheetCount
```

Copies selected worksheets from Source.xlsx into Target.xlsx using OfficeIMO workbook merge logic.

## PARAMETERS

### -CopyMode
Controls whether cross-workbook copies use package-level copy or value materialization.

```yaml
Type: ExcelWorksheetCopyMode
Parameter Sets: Path, Document
Aliases: None
Possible values: Values, Package

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Target workbook to update outside the DSL context.

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

### -Path
Target workbook path to create or update.

```yaml
Type: String
Parameter Sets: Path
Aliases: InputPath, FilePath, OutputPath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SheetNamePrefix
Prefix added to every imported worksheet name.

```yaml
Type: String
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SourceDocument
Optional source workbook object.

```yaml
Type: ExcelDocument
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SourcePath
Source workbook paths to merge into the target workbook.

```yaml
Type: String[]
Parameter Sets: Path, Document
Aliases: FullName, LiteralPath
Possible values:

Required: False
Position: 1
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -SourceSheet
Specific source worksheet names to import. Defaults to all source sheets.

```yaml
Type: String[]
Parameter Sets: Path, Document
Aliases: SheetName, Sheet, WorksheetName
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ValidationMode
Controls how invalid or duplicate destination sheet names are handled.

```yaml
Type: ExcelSheetNameValidationMode
Parameter Sets: Path, Document
Aliases: None
Possible values: None, Sanitize, Strict

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
- `System.String[]`

## OUTPUTS

- `OfficeIMO.Excel.ExcelWorkbookMergeResult`

## RELATED LINKS

- None
