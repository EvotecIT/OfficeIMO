---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Close-OfficeExcel
## SYNOPSIS
Closes a live Excel workbook without saving pending changes.

## SYNTAX
### __AllParameterSets
```powershell
Close-OfficeExcel [-Document] <ExcelDocument> [<CommonParameters>]
```

## DESCRIPTION
Use Save-OfficeExcel before closing to retain changes. Workbooks with SaveOnDispose are rejected so disposal cannot publish a file.

## EXAMPLES

### EXAMPLE 1
```powershell
$workbook | Save-OfficeExcel
            $workbook | Close-OfficeExcel
```


## PARAMETERS

### -Document
Explicit-persistence workbook to close.

```yaml
Type: ExcelDocument
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
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
