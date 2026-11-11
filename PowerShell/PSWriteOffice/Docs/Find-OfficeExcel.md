---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Find-OfficeExcel
## SYNOPSIS
Finds cells whose values match a pattern in a workbook, sheet or range.

## SYNTAX
### Document (Default)
```powershell
Find-OfficeExcel [-Pattern] <string> [-Sheet <OfficeSheetSelector[]>] [-Range <string>] [-SimpleMatch] [-CaseSensitive] [-Exact] [-Document <ExcelDocument>] [<CommonParameters>]
```

### Path
```powershell
Find-OfficeExcel [-Path] <string[]> [-Pattern] <string> [-Sheet <OfficeSheetSelector[]>] [-Range <string>] [-SimpleMatch] [-CaseSensitive] [-Exact] [<CommonParameters>]
```

### LiteralPath
```powershell
Find-OfficeExcel [-Pattern] <string> -LiteralPath <string[]> [-Sheet <OfficeSheetSelector[]>] [-Range <string>] [-SimpleMatch] [-CaseSensitive] [-Exact] [<CommonParameters>]
```

## DESCRIPTION
Finds cells whose values match a pattern in a workbook, sheet or range.

## EXAMPLES

### EXAMPLE 1
```powershell
Find-OfficeExcel -Path .\Reports\Inventory.xlsx -Pattern 'WEB01' -SimpleMatch
```

Lists every cell in the workbook that mentions the server WEB01.

### EXAMPLE 2
```powershell
$old = Import-OfficeCsv -Path .\Security\Retired-Certificates.csv
foreach ($certificate in $old) {
    Find-OfficeExcel -Path .\Security\Certificates.xlsx -Sheet Inventory -Pattern $certificate.Thumbprint -SimpleMatch -Exact
}
```

Looks up each retired certificate thumbprint on the inventory sheet and shows where it still appears.

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

### -Exact
Require an exact cell text match instead of substring matching.

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

### -Sheet
Worksheet name, one-based position or native worksheet object. Multi-sheet commands also accept wildcard names.

```yaml
Type: OfficeSheetSelector[]
Parameter Sets: Document, Path, LiteralPath
Aliases: WorksheetName
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
