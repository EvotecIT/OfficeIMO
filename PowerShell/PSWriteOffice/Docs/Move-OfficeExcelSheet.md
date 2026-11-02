---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Move-OfficeExcelSheet
## SYNOPSIS
Moves a worksheet to a one-based workbook position.

## SYNTAX
### Default (Default)
```powershell
Move-OfficeExcelSheet -Position <int> [-Sheet <OfficeSheetSelector>] [-PassThru] [-Document <ExcelDocument>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Uses the explicit or enclosing live workbook. Save it explicitly to persist the new order.

## EXAMPLES

### EXAMPLE 1
```powershell
$book | Move-OfficeExcelSheet -Sheet Summary -Position 1
            Save-OfficeExcel -Document $book
```


## PARAMETERS

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
Gets or sets whether to return the moved native worksheet.

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

### -Position
Gets or sets the one-based destination position.

```yaml
Type: Int32
Parameter Sets: Default
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Sheet
Gets or sets a worksheet name, one-based position or native worksheet.

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

- `OfficeIMO.Excel.ExcelSheet`

## RELATED LINKS

- None
