---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficeExcelFormula
## SYNOPSIS
Lists native formula records, optionally filtered by worksheet, volatile functions or external references.

## SYNTAX
### Document (Default)
```powershell
Get-OfficeExcelFormula [-Volatile] [-External] [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector[]>] [<CommonParameters>]
```

### Path
```powershell
Get-OfficeExcelFormula [-Path] <string[]> [-Volatile] [-External] [-Sheet <OfficeSheetSelector[]>] [<CommonParameters>]
```

### LiteralPath
```powershell
Get-OfficeExcelFormula -LiteralPath <string[]> [-Volatile] [-External] [-Sheet <OfficeSheetSelector[]>] [<CommonParameters>]
```

## DESCRIPTION
Lists native formula records, optionally filtered by worksheet, volatile functions or external references.

## EXAMPLES

### EXAMPLE 1
```powershell
Get-OfficeExcelFormula -Path .\Models\*.xlsx -Volatile | Format-Table SheetName, Address, Formula
```


### EXAMPLE 2
```powershell
Get-ChildItem -LiteralPath .\Models -Filter *.xlsx | Get-OfficeExcelFormula -External
```


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

### -External
Return only formulas referring to external workbooks.

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
Literal workbook paths. FileInfo.FullName and provider PSPath bind here.

```yaml
Type: String[]
Parameter Sets: LiteralPath
Aliases: FullName, PSPath
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -Path
Workbook paths, including filesystem wildcard patterns.

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
Worksheet names, wildcard patterns, one-based positions or native worksheet objects.

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

### -Volatile
Return only formulas containing volatile functions.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.String[]`
- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `OfficeIMO.Excel.ExcelFormulaInfo`

## RELATED LINKS

- None
