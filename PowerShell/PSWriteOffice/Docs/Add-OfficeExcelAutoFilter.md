---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeExcelAutoFilter
## SYNOPSIS
Adds an AutoFilter to the current worksheet.

## SYNTAX
### Context (Default)
```powershell
Add-OfficeExcelAutoFilter [-Range] <string> [-Criteria <hashtable>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficeExcelAutoFilter [-Range] <string> -Document <ExcelDocument> [-Sheet <string>] [-SheetIndex <Int32>] [-Criteria <hashtable>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Optional filter criteria can be supplied per column index (0-based within the range).

## EXAMPLES

### EXAMPLE 1
```powershell
PS> ExcelSheet 'Data' { Add-OfficeExcelAutoFilter -Range 'A1:D200' }
```

Enables filter dropdowns on the range.

### EXAMPLE 2
```powershell
PS> Add-OfficeExcelAutoFilter -Range 'A1:D200' -Criteria @{ 2 = 'Open','Hold' }
```

Filters the third column (0-based within the range) to Open/Hold.

## PARAMETERS

### -Criteria
Optional criteria per column index (0-based within the range).

```yaml
Type: Hashtable
Parameter Sets: Context, Document
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
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -PassThru
Emit the object created or changed by the command.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Range
A1 range to apply AutoFilter.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Sheet
Worksheet name when using Document.

```yaml
Type: String
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SheetIndex
Worksheet index (0-based) when using Document.

```yaml
Type: Int32
Parameter Sets: Document
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
