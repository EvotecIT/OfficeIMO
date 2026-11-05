---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Invoke-OfficeExcelAutoFit
## SYNOPSIS
Fits column widths and row heights to their content.

## SYNTAX
### Default (Default)
```powershell
Invoke-OfficeExcelAutoFit [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-Column <Object[]>] [-Row <Object[]>] [<CommonParameters>]
```

## DESCRIPTION
Fits column widths and row heights to their content.

## EXAMPLES

### EXAMPLE 1
```powershell
Open-OfficeExcel -Path .\Reports\Inventory.xlsx {
    ExcelSheet 'Servers' {
        ExcelAutoFit
    }
}
```

Fits all columns and rows of the current sheet to their content.

### EXAMPLE 2
```powershell
Open-OfficeExcel -Path .\Reports\Changes.xlsx {
    ExcelSheet 'Changes' {
        ExcelColumn F -Values (Import-OfficeCsv -Path .\Data\Change-Notes.csv).Note -StartRow 2
        ExcelAutoFit -Column E, F
    }
}
```

Writes long notes into column F and refits only columns E and F, leaving the widths of the rest as designed.

## PARAMETERS

### -Column
Auto-fit specific column indexes (1-based).

```yaml
Type: Object[]
Parameter Sets: Default
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
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Row
Auto-fit specific row indexes (1-based).

```yaml
Type: Object[]
Parameter Sets: Default
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
Type: OfficeSheetSelector
Parameter Sets: Default
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

- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `None`

## RELATED LINKS

- None
