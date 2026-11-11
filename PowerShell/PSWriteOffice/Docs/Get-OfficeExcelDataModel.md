---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficeExcelDataModel
## SYNOPSIS
Reports connection, query-table, data-model and external-link parts in a workbook.

## SYNTAX
### Document (Default)
```powershell
Get-OfficeExcelDataModel [-Document <ExcelDocument>] [<CommonParameters>]
```

### Path
```powershell
Get-OfficeExcelDataModel [-Path] <string[]> [<CommonParameters>]
```

### LiteralPath
```powershell
Get-OfficeExcelDataModel -LiteralPath <string[]> [<CommonParameters>]
```

## DESCRIPTION
Reports connection, query-table, data-model and external-link parts in a workbook.

## EXAMPLES

### EXAMPLE 1
```powershell
Get-OfficeExcelDataModel -Path .\Reports\Live-Sales.xlsx
```

Lists connections, query tables, data-model parts and external links in the workbook.

### EXAMPLE 2
```powershell
Get-ChildItem -Path \\files\Finance -Filter *.xlsx -Recurse |
                Get-OfficeExcelDataModel |
                Format-List
```

Reports the data model of every workbook on the finance share, to find files that depend on databases or other workbooks.

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

### -Path
Workbook paths resolved with filesystem wildcard support.

```yaml
Type: String[]
Parameter Sets: Path
Aliases: FilePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
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
