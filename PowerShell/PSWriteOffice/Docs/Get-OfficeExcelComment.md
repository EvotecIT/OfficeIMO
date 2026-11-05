---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficeExcelComment
## SYNOPSIS
Gets notes and threaded comments that match a filter.

## SYNTAX
### Document (Default)
```powershell
Get-OfficeExcelComment [-Sheet <OfficeSheetSelector[]>] [-Range <string>] [-Author <string>] [-TextContains <string>] [-Kind <ExcelCommentKind>] [-Document <ExcelDocument>] [<CommonParameters>]
```

### Path
```powershell
Get-OfficeExcelComment [-Path] <string[]> [-Sheet <OfficeSheetSelector[]>] [-Range <string>] [-Author <string>] [-TextContains <string>] [-Kind <ExcelCommentKind>] [<CommonParameters>]
```

### LiteralPath
```powershell
Get-OfficeExcelComment -LiteralPath <string[]> [-Sheet <OfficeSheetSelector[]>] [-Range <string>] [-Author <string>] [-TextContains <string>] [-Kind <ExcelCommentKind>] [<CommonParameters>]
```

## DESCRIPTION
Gets notes and threaded comments that match a filter.

## EXAMPLES

### EXAMPLE 1
```powershell
Get-OfficeExcelComment -Path .\Finance\Forecast-2027.xlsx
```

Shows every note and threaded comment in the workbook.

### EXAMPLE 2
```powershell
Get-OfficeExcelComment -Path .\Finance\Forecast-2027.xlsx -Kind Threaded -Author 'Anna Kowalska' |
                Export-OfficeExcel -Path .\Reviews\Forecast-Comments.xlsx -Title 'Open review comments'
```

Finds the threaded comments left by the finance reviewer and exports them as a review list.

## PARAMETERS

### -Author
Comment author to match, ignoring case.

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

### -Document
Caller-owned workbook; omission uses the enclosing Excel content block.

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

### -Kind
Metadata kind to add.

```yaml
Type: ExcelCommentKind
Parameter Sets: Document, Path, LiteralPath
Aliases: None
Possible values: All, Note, Threaded

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LiteralPath
Literal workbook paths, including FileInfo.FullName and provider PSPath.

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

### -Path
Workbook paths resolved with filesystem wildcard support.

```yaml
Type: String[]
Parameter Sets: Path
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Range
A1 cell range on the selected worksheet.

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

### -Sheet
Worksheet name, one-based position or native worksheet object. Multi-sheet commands also accept wildcard names.

```yaml
Type: OfficeSheetSelector[]
Parameter Sets: Document, Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TextContains
Text fragment to match, ignoring case.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.String[]`
- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `None`

## RELATED LINKS

- None
