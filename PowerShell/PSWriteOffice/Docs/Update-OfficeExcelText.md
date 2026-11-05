---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Update-OfficeExcelText
## SYNOPSIS
Replaces text in cell values across a workbook, sheet or range.

## SYNTAX
### Document (Default)
```powershell
Update-OfficeExcelText [-Pattern] <string> [-Replacement] <string> [-Sheet <OfficeSheetSelector[]>] [-Range <string>] [-SimpleMatch] [-CaseSensitive] [-Show] [-PassThru] [-Document <ExcelDocument>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Path
```powershell
Update-OfficeExcelText [-Path] <string[]> [-Pattern] <string> [-Replacement] <string> [-Sheet <OfficeSheetSelector[]>] [-Range <string>] [-SimpleMatch] [-CaseSensitive] [-Show] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPath
```powershell
Update-OfficeExcelText [-Pattern] <string> [-Replacement] <string> -LiteralPath <string[]> [-Sheet <OfficeSheetSelector[]>] [-Range <string>] [-SimpleMatch] [-CaseSensitive] [-Show] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Replaces text in cell values across a workbook, sheet or range.

## EXAMPLES

### EXAMPLE 1
```powershell
Update-OfficeExcelText -Path .\Templates\Price-List.xlsx -Pattern 'Fabrikam Ltd' -Replacement 'Northwind Traders' -SimpleMatch
```

Replaces the old company name in every cell of the workbook and saves it.

### EXAMPLE 2
```powershell
$changed = Update-OfficeExcelText -Path .\Reports\Inventory.xlsx -Sheet Servers, Services -Pattern '^SRV-(\d+)$' -Replacement 'NW-APP-$1' -PassThru
"Renamed $changed cells"
```

Uses a regular expression to rename the old SRV hosts to the new naming scheme on the inventory sheets and reports how many cells changed.

### EXAMPLE 3
```powershell
Open-OfficeExcel -Path .\Templates\Quote.xlsx {
    Update-OfficeExcelText -Pattern '{{Customer}}' -Replacement 'Northwind Traders' -SimpleMatch
    Update-OfficeExcelText -Pattern '{{Date}}' -Replacement (Get-Date -Format 'yyyy-MM-dd') -SimpleMatch
}
```

Fills a template and replaces its placeholder text in the same Open-OfficeExcel block.

## PARAMETERS

### -CaseSensitive
Matches names and text using ordinal case sensitivity. Omission ignores case.

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

### -PassThru
Emits the native result after the operation.

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

### -Pattern
Date or DateTime number format pattern.

```yaml
Type: String
Parameter Sets: Document, Path, LiteralPath
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
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

### -Replacement
Replacement text or regular-expression replacement pattern.

```yaml
Type: String
Parameter Sets: Document, Path, LiteralPath
Aliases: None
Possible values:

Required: True
Position: 2
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

### -Show
Open the saved file in its registered application after a local block completes.

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

### -SimpleMatch
Uses literal text matching instead of regular expressions.

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

- `None`

## RELATED LINKS

- None
