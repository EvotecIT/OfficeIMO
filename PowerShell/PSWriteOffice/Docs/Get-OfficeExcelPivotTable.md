---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficeExcelPivotTable
## SYNOPSIS
Lists pivot tables in a workbook.

## SYNTAX
### Document (Default)
```powershell
Get-OfficeExcelPivotTable [-Name <string>] [-Sheet <OfficeSheetSelector>] [-Document <ExcelDocument>] [<CommonParameters>]
```

### Path
```powershell
Get-OfficeExcelPivotTable [-Path] <string[]> [-Name <string>] [-Sheet <OfficeSheetSelector>] [<CommonParameters>]
```

### LiteralPath
```powershell
Get-OfficeExcelPivotTable -LiteralPath <string[]> [-Name <string>] [-Sheet <OfficeSheetSelector>] [<CommonParameters>]
```

## DESCRIPTION
Lists pivot tables in a workbook.

## EXAMPLES

### EXAMPLE 1
```powershell
Get-OfficeExcelPivotTable -Path .\Reports\Sales-2026.xlsx
```

Shows every pivot table in the workbook with its location and source.

### EXAMPLE 2
```powershell
Get-OfficeExcelPivotTable -Path .\Reports\Sales-2026.xlsx |
                Select-Object Name, SheetName, Location, SourceSheet, SourceRange
```

Lists the source range of each pivot table, to confirm they still point at the current data after the sheet was rebuilt.

## PARAMETERS

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

### -Name
Optional pivot table name filter.

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
