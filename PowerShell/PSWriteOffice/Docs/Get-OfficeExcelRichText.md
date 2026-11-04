---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficeExcelRichText
## SYNOPSIS
Gets the rich-text runs of a cell.

## SYNTAX
### Document (Default)
```powershell
Get-OfficeExcelRichText -Address <string> [-Sheet <OfficeSheetSelector>] [-Document <ExcelDocument>] [<CommonParameters>]
```

### Path
```powershell
Get-OfficeExcelRichText [-Path] <string[]> -Address <string> [-Sheet <OfficeSheetSelector>] [<CommonParameters>]
```

### LiteralPath
```powershell
Get-OfficeExcelRichText -Address <string> -LiteralPath <string[]> [-Sheet <OfficeSheetSelector>] [<CommonParameters>]
```

## DESCRIPTION
Gets the rich-text runs of a cell.

## EXAMPLES

### EXAMPLE 1
```powershell
Get-OfficeExcelRichText -Path .\Reports\Migration.xlsx -Sheet Overview -Address B3
```

Returns each formatted run of B3 on the Overview sheet.

### EXAMPLE 2
```powershell
$book = Open-OfficeExcel -Path .\Reports\Status.xlsx -ReadOnly
Get-OfficeExcelRichText -Document $book -Sheet Summary -Address B2 | Where-Object Bold
Close-OfficeExcel -Document $book
```

Reads the runs of a status cell from an open workbook and keeps only the bold text.

## PARAMETERS

### -Address
A1 address of the target cell.

```yaml
Type: String
Parameter Sets: Document, Path, LiteralPath
Aliases: None
Possible values:

Required: True
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

### -Sheet
Worksheet name, one-based position or native worksheet object. Multi-sheet commands also accept wildcard names.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.String[]`
- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `None`

## RELATED LINKS

- None
