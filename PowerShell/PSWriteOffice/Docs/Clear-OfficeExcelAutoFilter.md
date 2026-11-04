---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Clear-OfficeExcelAutoFilter
## SYNOPSIS
Clears any AutoFilter on the current worksheet.

## SYNTAX
### Default (Default)
```powershell
Clear-OfficeExcelAutoFilter [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Clears any AutoFilter on the current worksheet.

## EXAMPLES

### EXAMPLE 1
```powershell
Open-OfficeExcel -Path .\Reports\Helpdesk.xlsx {
    ExcelSheet 'Tickets' {
        Clear-OfficeExcelAutoFilter
    }
}
```

Removes the filter buttons and any filter conditions from the Tickets sheet.

### EXAMPLE 2
```powershell
$book = Open-OfficeExcel -Path .\Reports\Q3-Sales.xlsx
foreach ($sheet in 'Orders', 'Returns', 'Customers') {
    Clear-OfficeExcelAutoFilter -Document $book -Sheet $sheet
}
Save-OfficeExcel -Document $book -Path .\Outbox\Q3-Sales.xlsx -PassThru | Close-OfficeExcel
```

Removes the filters left on the working sheets so recipients see every row.

## PARAMETERS

### -Document
Workbook to operate on outside the DSL context.

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
Emit the object created or changed by the command.

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

### -Sheet
Worksheet name when using Document.

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
