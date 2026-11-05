---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Remove-OfficeExcelComment
## SYNOPSIS
Removes notes from a cell or all notes that match a filter.

## SYNTAX
### Default (Default)
```powershell
Remove-OfficeExcelComment [[-Range] <string>] [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector[]>] [-Author <string>] [-TextContains <string>] [-All] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Removes notes from a cell or all notes that match a filter.

## EXAMPLES

### EXAMPLE 1
```powershell
Open-OfficeExcel -Path .\Finance\Budget-2027.xlsx {
    Remove-OfficeExcelComment B4 -Sheet Assumptions
}
```

Deletes the note on B4 of the Assumptions sheet.

### EXAMPLE 2
```powershell
Open-OfficeExcel -Path .\Reports\Q3-Sales.xlsx {
    $removed = Remove-OfficeExcelComment -Author 'Review Bot' -PassThru
    Write-Verbose "Removed $removed review notes"
}
```

Removes every note left by the review bot across all sheets and reports how many were removed.

## PARAMETERS

### -All
Explicitly selects every item when no filter is supplied.

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

### -Author
Comment author to match, ignoring case.

```yaml
Type: String
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

### -PassThru
Emits the native result after the operation.

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

### -Range
A1 cell range on the selected worksheet.

```yaml
Type: String
Parameter Sets: Default
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
Type: OfficeSheetSelector[]
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TextContains
Text fragment to match, ignoring case.

```yaml
Type: String
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
