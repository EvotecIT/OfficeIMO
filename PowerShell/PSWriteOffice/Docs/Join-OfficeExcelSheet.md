---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Join-OfficeExcelSheet
## SYNOPSIS
Appends or merges rows from one worksheet into another.

## SYNTAX
### Context (Default)
```powershell
Join-OfficeExcelSheet -SourceSheet <string> [-TargetSheet <string>] [-TargetSheetIndex <Int32>] [-SourceDocument <ExcelDocument>] [-SourcePath <string>] [-SourceRange <string>] [-TargetStartRow <Int32>] [-TargetStartColumn <Int32>] [-NoSourceHeader] [-IncludeSourceHeader] [-MatchColumnsByHeader] [-TargetHeaderRow <Int32>] [-BlankRowsBefore <int>] [-OverwriteExistingCells] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Path
```powershell
Join-OfficeExcelSheet [-Path] <string> -SourceSheet <string> [-TargetSheet <string>] [-TargetSheetIndex <Int32>] [-SourceDocument <ExcelDocument>] [-SourcePath <string>] [-SourceRange <string>] [-TargetStartRow <Int32>] [-TargetStartColumn <Int32>] [-NoSourceHeader] [-IncludeSourceHeader] [-MatchColumnsByHeader] [-TargetHeaderRow <Int32>] [-BlankRowsBefore <int>] [-OverwriteExistingCells] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
Join-OfficeExcelSheet -Document <ExcelDocument> -SourceSheet <string> [-TargetSheet <string>] [-TargetSheetIndex <Int32>] [-SourceDocument <ExcelDocument>] [-SourcePath <string>] [-SourceRange <string>] [-TargetStartRow <Int32>] [-TargetStartColumn <Int32>] [-NoSourceHeader] [-IncludeSourceHeader] [-MatchColumnsByHeader] [-TargetHeaderRow <Int32>] [-BlankRowsBefore <int>] [-OverwriteExistingCells] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Appends or merges rows from one worksheet into another.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $result = Join-OfficeExcelSheet -Path .\Report.xlsx -TargetSheet Combined -SourceSheet Data -MatchColumnsByHeader -BlankRowsBefore 1
$result |
    Select-Object -Property SourceSheet, TargetSheet, RowsCopied, ColumnsCopied
```

Copies rows from Data into Combined, aligns columns by header, and returns the merge result.

## PARAMETERS

### -BlankRowsBefore
Blank rows to leave before appended data.

```yaml
Type: Int32
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

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

### -IncludeSourceHeader
Include the source header row in copied rows.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MatchColumnsByHeader
Match source columns to target columns by header text.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoSourceHeader
Treat the first source row as data instead of a header row.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OverwriteExistingCells
Allow copied values to replace existing target cells.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Target workbook path to update.

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

### -SourceDocument
Optional source workbook object for cross-workbook joins.

```yaml
Type: ExcelDocument
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SourcePath
Optional source workbook path for cross-workbook joins.

```yaml
Type: String
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SourceRange
Source A1 range to copy. Defaults to the source used range.

```yaml
Type: String
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SourceSheet
Source worksheet name.

```yaml
Type: String
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TargetHeaderRow
1-based target header row when matching columns by header.

```yaml
Type: Int32
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TargetSheet
Target worksheet name. Defaults to the current sheet inside an ExcelSheet block.

```yaml
Type: String
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TargetSheetIndex
Target worksheet index when using a workbook object or path.

```yaml
Type: Int32
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TargetStartColumn
1-based target start column. Defaults to the source range start column.

```yaml
Type: Int32
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TargetStartRow
1-based target start row. Defaults to appending after the target used range.

```yaml
Type: Int32
Parameter Sets: Context, Path, Document
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

- `OfficeIMO.Excel.ExcelWorksheetMergeResult`

## RELATED LINKS

- None
