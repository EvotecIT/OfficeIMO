---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficeExcelPageBreak
## SYNOPSIS
Adds manual print page breaks to an open worksheet.

## SYNTAX
### Default (Default)
```powershell
Add-OfficeExcelPageBreak [-Sheet <OfficeSheetSelector>] [-Row <int[]>] [-Column <int[]>] [-PassThru] [-Document <ExcelDocument>] [<CommonParameters>]
```

## DESCRIPTION
Adds manual print page breaks to an open worksheet.

## EXAMPLES

### EXAMPLE 1
```powershell
ExcelSheet 'Data' { ExcelPageBreak -Row 25,50 -Column 8 }
```


### EXAMPLE 2
```powershell
Open-OfficeExcel -Path .\Reports\Inventory.xlsx {
    ExcelSheet 'Servers' {
        ExcelPageBreak -Row 41
    }
}
```

Starts a new printed page at row 41.

### EXAMPLE 3
```powershell
$sites = Import-OfficeExcel -Path .\Reports\Inventory.xlsx -Sheet Servers
$firstRows = @()
$rowNumber = 2
foreach ($group in $sites | Group-Object Site) {
    if ($rowNumber -gt 2) { $firstRows += $rowNumber }
    $rowNumber += $group.Count
}
Open-OfficeExcel -Path .\Reports\Inventory.xlsx {
    ExcelSheet 'Servers' {
        ExcelPageBreak -Row $firstRows
    }
}
```

Adds a page break above the first row of each site block so every site starts on a new page.

## PARAMETERS

### -Column
One-based columns after which printing starts a new page.

```yaml
Type: Int32[]
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
Gets or sets an explicit live document; omission uses the enclosing content block.

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

### -PassThru
Emit the worksheet's page-break snapshots after adding the selected breaks.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Row
One-based rows after which printing starts a new page.

```yaml
Type: Int32[]
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
Worksheet name, one-based position or native worksheet; omission uses the current sheet.

```yaml
Type: OfficeSheetSelector
Parameter Sets: Default
Aliases: WorksheetName
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

- `PSWriteOffice.Models.Excel.ExcelPageBreakInfo`: A manual print page-break snapshot that remains readable after its source workbook is closed.

## RELATED LINKS

- None
