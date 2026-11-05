---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficeExcelCompareOptions
## SYNOPSIS
Creates workbook comparison settings for Compare-OfficeDocument.

## SYNTAX
### Default (Default)
```powershell
New-OfficeExcelCompareOptions [-Exclude <ExcelCompareArea[]>] [<CommonParameters>]
```

## DESCRIPTION
Creates workbook comparison settings for Compare-OfficeDocument.

## EXAMPLES

### EXAMPLE 1
```powershell
Compare-OfficeDocument -ReferencePath .\Finance\Budget-v1.xlsx -DifferencePath .\Finance\Budget-v2.xlsx -Options (New-OfficeExcelCompareOptions -Exclude CellStyles)
```

Compares two versions of a workbook while ignoring cell style changes.

### EXAMPLE 2
```powershell
$compare = New-OfficeExcelCompareOptions -Exclude CellStyles, Comments
Compare-OfficeDocument -ReferencePath .\Pricing\Price-List-2026-09.xlsx -DifferencePath .\Pricing\Price-List-2026-10.xlsx -Options $compare -OutputPath .\Pricing\Price-Changes-October.xlsx
```

Compares last month's and this month's price list, ignoring styles and comments, and saves the comparison report.

## PARAMETERS

### -Exclude
Worksheet names or wildcard patterns omitted from the navigation list.

```yaml
Type: ExcelCompareArea[]
Parameter Sets: Default
Aliases: None
Possible values: Cells, CellStyles, NamedRanges, Tables, WorksheetMetadata, Comments

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `None`

## RELATED LINKS

- None
