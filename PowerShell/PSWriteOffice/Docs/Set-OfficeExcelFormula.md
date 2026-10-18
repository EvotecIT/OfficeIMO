---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeExcelFormula
## SYNOPSIS
Sets a formula in a worksheet cell.

## SYNTAX
### Coordinates
```powershell
Set-OfficeExcelFormula -Formula <string> [-Row <Int32>] [-Column <Int32>] [-PassThru] [<CommonParameters>]
```

### Address
```powershell
Set-OfficeExcelFormula -Formula <string> [-Address <string>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Supports A1 addresses or row/column coordinates in the Excel DSL.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> ExcelSheet 'Data' { Set-OfficeExcelFormula -Address 'C2' -Formula 'SUM(A2:B2)' }
```

Stores the formula in cell C2.

## PARAMETERS

### -Address
A1-style cell address (e.g., A1, C5).

```yaml
Type: String
Parameter Sets: Address
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Column
1-based column index.

```yaml
Type: Int32
Parameter Sets: Coordinates
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Formula
Formula text (without leading =).

```yaml
Type: String
Parameter Sets: Coordinates, Address
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the object created or changed by the command.

```yaml
Type: SwitchParameter
Parameter Sets: Coordinates, Address
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Row
1-based row index.

```yaml
Type: Int32
Parameter Sets: Coordinates
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

- `None`

## OUTPUTS

- `None`

## RELATED LINKS

- None
