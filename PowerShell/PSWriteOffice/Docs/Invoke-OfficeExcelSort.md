---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Invoke-OfficeExcelSort
## SYNOPSIS
Sorts the used range on the current worksheet.

## SYNTAX
### Header (Default)
```powershell
Invoke-OfficeExcelSort [-Header] <string> [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-Descending] [-PassThru] [<CommonParameters>]
```

### Order
```powershell
Invoke-OfficeExcelSort -Order <ordered> [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Sorts the used range on the current worksheet.

## EXAMPLES

### EXAMPLE 1
```powershell
Open-OfficeExcel -Path .\Reports\Inventory.xlsx {
    ExcelSheet 'Servers' {
        ExcelSort Name
    }
}
```

Sorts the used range of the Servers sheet by the Name header.

### EXAMPLE 2
```powershell
Open-OfficeExcel -Path .\Reports\Helpdesk.xlsx {
    ExcelSheet 'Tickets' {
        ExcelSort -Order ([ordered]@{ Priority = $true; AgeDays = $false })
    }
}
```

Sorts tickets by priority ascending and then by age with the oldest first.

## PARAMETERS

### -Descending
Sort descending (single-column sort).

```yaml
Type: SwitchParameter
Parameter Sets: Header
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Workbook to operate on outside the DSL context.

```yaml
Type: ExcelDocument
Parameter Sets: Header, Order
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Header
Header to sort by (single-column sort).

```yaml
Type: String
Parameter Sets: Header
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Order
Ordered dictionary of header => ascending (true/false).

```yaml
Type: OrderedDictionary
Parameter Sets: Order
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the worksheet after sorting.

```yaml
Type: SwitchParameter
Parameter Sets: Header, Order
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
Parameter Sets: Header, Order
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
