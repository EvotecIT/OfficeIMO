---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Set-OfficeExcelAutoFilter
## SYNOPSIS
Adds an AutoFilter to a range and applies filter conditions.

## SYNTAX
### Range (Default)
```powershell
Set-OfficeExcelAutoFilter [-Range] <string> [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-PassThru] [-Criteria <hashtable>] [<CommonParameters>]
```

### Condition
```powershell
Set-OfficeExcelAutoFilter [-Header] <string> [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-PassThru] [-Range <string>] [-Value <string[]>] [-Contains <string>] [-DoesNotContain <string>] [-StartsWith <string>] [-EndsWith <string>] [-GreaterThanOrEqual <Double>] [-LessThanOrEqual <Double>] [-NotEqual <Double>] [-Between <double[]>] [-NotBetween <double[]>] [<CommonParameters>]
```

## DESCRIPTION
Adds an AutoFilter to a range and applies filter conditions.

## EXAMPLES

### EXAMPLE 1
```powershell
Open-OfficeExcel -Path .\Reports\Inventory.xlsx {
    ExcelSheet 'Servers' {
        ExcelAutoFilter A1:F500
    }
}
```

Adds AutoFilter drop-downs to the header row of A1:F500.

### EXAMPLE 2
```powershell
Open-OfficeExcel -Path .\Reports\Helpdesk.xlsx {
    ExcelSheet 'Tickets' {
        ExcelAutoFilter -Header Status -Value 'Open', 'On hold'
        ExcelAutoFilter -Header AgeDays -GreaterThanOrEqual 3
    }
}
```

Filters the ticket list to the Open and On hold statuses with at least three days of age, so the sheet opens on the work that needs attention.

### EXAMPLE 3
```powershell
Open-OfficeExcel -Path .\Reports\Helpdesk.xlsx {
    ExcelSheet 'Tickets' {
        ExcelAutoFilter A1:H900 -Criteria @{ Status = 'Open', 'On hold'; Region = 'EMEA' }
    }
}
```

Adds filter buttons and immediately shows only open EMEA tickets.

## PARAMETERS

### -Between
Inclusive numeric range condition. Provide exactly two values: minimum, maximum.

```yaml
Type: Double[]
Parameter Sets: Condition
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Contains
Text that the column value must contain.

```yaml
Type: String
Parameter Sets: Condition
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Criteria
Maps header names to filter values, value lists or comparison criteria.

```yaml
Type: Hashtable
Parameter Sets: Range
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
Parameter Sets: Range, Condition
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -DoesNotContain
Text that the column value must not contain.

```yaml
Type: String
Parameter Sets: Condition
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -EndsWith
Text that the column value must end with.

```yaml
Type: String
Parameter Sets: Condition
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -GreaterThanOrEqual
Numeric greater-than-or-equal condition.

```yaml
Type: Double
Parameter Sets: Condition
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Header
Header text identifying the target column.

```yaml
Type: String
Parameter Sets: Condition
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LessThanOrEqual
Numeric less-than-or-equal condition.

```yaml
Type: Double
Parameter Sets: Condition
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NotBetween
Outside numeric range condition. Provide exactly two values: minimum, maximum.

```yaml
Type: Double[]
Parameter Sets: Condition
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NotEqual
Numeric not-equal condition.

```yaml
Type: Double
Parameter Sets: Condition
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emits the native result after the operation.

```yaml
Type: SwitchParameter
Parameter Sets: Range, Condition
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Range
A1 cell range on the selected worksheet.

```yaml
Type: String
Parameter Sets: Range, Condition
Aliases: None
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Sheet
Worksheet name, one-based position or native worksheet object. Multi-sheet commands also accept wildcard names.

```yaml
Type: OfficeSheetSelector
Parameter Sets: Range, Condition
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -StartsWith
Text that the column value must start with.

```yaml
Type: String
Parameter Sets: Condition
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Value
Value or values consumed by the selected operation.

```yaml
Type: String[]
Parameter Sets: Condition
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
