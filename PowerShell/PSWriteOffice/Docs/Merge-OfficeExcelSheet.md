---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Merge-OfficeExcelSheet
## SYNOPSIS
Appends worksheet rows or merges a source range into a target worksheet.

## SYNTAX
### Document (Default)
```powershell
Merge-OfficeExcelSheet -SourceSheet <OfficeSheetSelector> [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-Source <Object>] [-SourceRange <string>] [-Address <string>] [-NoSourceHeader] [-IncludeSourceHeader] [-MatchColumnsByHeader] [-HeaderRow <Int32>] [-BlankRowsBefore <int>] [-OverwriteExistingCells] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Path
```powershell
Merge-OfficeExcelSheet [-Path] <string[]> -SourceSheet <OfficeSheetSelector> [-Sheet <OfficeSheetSelector>] [-Source <Object>] [-SourceRange <string>] [-Address <string>] [-NoSourceHeader] [-IncludeSourceHeader] [-MatchColumnsByHeader] [-HeaderRow <Int32>] [-BlankRowsBefore <int>] [-OverwriteExistingCells] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPath
```powershell
Merge-OfficeExcelSheet -LiteralPath <string[]> -SourceSheet <OfficeSheetSelector> [-Sheet <OfficeSheetSelector>] [-Source <Object>] [-SourceRange <string>] [-Address <string>] [-NoSourceHeader] [-IncludeSourceHeader] [-MatchColumnsByHeader] [-HeaderRow <Int32>] [-BlankRowsBefore <int>] [-OverwriteExistingCells] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Source accepts a native ExcelDocument, a workbook path or FileInfo. Omit Source to copy rows
within the target workbook. File targets are saved in place; live documents remain caller-owned
and require an explicit save. The result reports the actual copied rows and columns.

## EXAMPLES

### EXAMPLE 1
```powershell
Merge-OfficeExcelSheet -Path ./Sales.xlsx -Sheet Year -SourceSheet October -MatchColumnsByHeader
```


### EXAMPLE 2
```powershell
Get-Item -LiteralPath './Combined[1].xlsx' | Merge-OfficeExcelSheet -Sheet Combined -Source (Get-Item -LiteralPath './Branch[1].xlsx') -SourceSheet Sales -Address B5 -HeaderRow 4 -MatchColumnsByHeader
```


## PARAMETERS

### -Address
Gets or sets the target start cell; omission appends after the used range.

```yaml
Type: String
Parameter Sets: Document, Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -BlankRowsBefore
Gets or sets blank rows before appended data; an explicit Address takes precedence.

```yaml
Type: Int32
Parameter Sets: Document, Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Gets or sets the live target workbook; omission uses the enclosing worksheet scope.

```yaml
Type: ExcelDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -HeaderRow
Gets or sets the one-based target header row used for matching columns.

```yaml
Type: Int32
Parameter Sets: Document, Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IncludeSourceHeader
Gets or sets whether to copy the source header along with its data rows.

```yaml
Type: SwitchParameter
Parameter Sets: Document, Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LiteralPath
Gets or sets literal target paths, including FileInfo.FullName and provider PSPath.

```yaml
Type: String[]
Parameter Sets: LiteralPath
Aliases: PSPath, FullName
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -MatchColumnsByHeader
Gets or sets whether columns are aligned by header text rather than position.

```yaml
Type: SwitchParameter
Parameter Sets: Document, Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoSourceHeader
Gets or sets whether the first source row is data rather than a header.

```yaml
Type: SwitchParameter
Parameter Sets: Document, Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OverwriteExistingCells
Gets or sets whether source values may replace existing target cells.

```yaml
Type: SwitchParameter
Parameter Sets: Document, Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Gets or sets target workbook paths with filesystem wildcard expansion.

```yaml
Type: String[]
Parameter Sets: Path
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -Sheet
Gets or sets the target worksheet name, one-based position or native worksheet.

```yaml
Type: OfficeSheetSelector
Parameter Sets: Document, Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Source
Gets or sets another workbook as a native ExcelDocument, path or literal FileInfo.

```yaml
Type: Object
Parameter Sets: Document, Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SourceRange
Gets or sets the source A1 range; omission copies its used range.

```yaml
Type: String
Parameter Sets: Document, Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SourceSheet
Gets or sets the source worksheet name, one-based position or native worksheet.

```yaml
Type: OfficeSheetSelector
Parameter Sets: Document, Path, LiteralPath
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.String[]`
- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `OfficeIMO.Excel.ExcelWorksheetMergeResult`

## RELATED LINKS

- None
