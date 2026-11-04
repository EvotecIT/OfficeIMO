---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Copy-OfficeExcelSheet
## SYNOPSIS
Copies selected worksheets within a workbook or from another workbook.

## SYNTAX
### Document (Default)
```powershell
Copy-OfficeExcelSheet [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector[]>] [-Name <string>] [-NamePrefix <string>] [-Source <Object>] [-ValidationMode <ExcelSheetNameValidationMode>] [-CopyMode <ExcelWorksheetCopyMode>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Path
```powershell
Copy-OfficeExcelSheet [-Path] <string[]> [-Sheet <OfficeSheetSelector[]>] [-Name <string>] [-NamePrefix <string>] [-Source <Object>] [-ValidationMode <ExcelSheetNameValidationMode>] [-CopyMode <ExcelWorksheetCopyMode>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPath
```powershell
Copy-OfficeExcelSheet -LiteralPath <string[]> [-Sheet <OfficeSheetSelector[]>] [-Name <string>] [-NamePrefix <string>] [-Source <Object>] [-ValidationMode <ExcelSheetNameValidationMode>] [-CopyMode <ExcelWorksheetCopyMode>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Source accepts a native ExcelDocument, a workbook path or literal FileInfo. Sheet accepts names,
wildcard patterns, one-based positions and native worksheets. Omit Sheet to copy all source sheets,
or the current worksheet inside a worksheet content block. Native package copy preserves references
between selected sheets. File targets are saved in place; live documents require an explicit save.

PassThru emits native worksheets. For file input, their workbook remains open: close
the returned worksheet's Document with Close-OfficeExcel when finished.

## EXAMPLES

### EXAMPLE 1
```powershell
Copy-OfficeExcelSheet -Document $book -Sheet Template -Name October
```


### EXAMPLE 2
```powershell
Copy-OfficeExcelSheet -LiteralPath ./Combined.xlsx -Source ./Region.xlsx -Sheet 'Sales*' -NamePrefix 'Region '
```


## PARAMETERS

### -CopyMode
Gets or sets native package copying or explicit value-only materialization.

```yaml
Type: ExcelWorksheetCopyMode
Parameter Sets: Document, Path, LiteralPath
Aliases: None
Possible values: Values, Package

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Gets or sets a live target workbook; omission uses the enclosing workbook scope.

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

### -Name
Gets or sets a new name when copying exactly one worksheet.

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

### -NamePrefix
Gets or sets a prefix for each copied worksheet name.

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

### -PassThru
Gets or sets whether to return live copied worksheets; their workbook is caller-owned.

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
Gets or sets source worksheet names or patterns, one-based positions or native worksheets.

```yaml
Type: OfficeSheetSelector[]
Parameter Sets: Document, Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
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

### -ValidationMode
Gets or sets how invalid or duplicate destination names are handled.

```yaml
Type: ExcelSheetNameValidationMode
Parameter Sets: Document, Path, LiteralPath
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

- `System.String[]`
- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `OfficeIMO.Excel.ExcelSheet`

## RELATED LINKS

- None
