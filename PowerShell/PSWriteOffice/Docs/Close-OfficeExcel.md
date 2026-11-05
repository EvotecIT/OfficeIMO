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
### Default (Default)
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


### EXAMPLE 2
```powershell
$book = Open-OfficeExcel -Path .\Reports\Q3-Sales.xlsx -ReadOnly
Close-OfficeExcel -Document $book
```

Releases a workbook that was only read; nothing is written to disk.

### EXAMPLE 3
```powershell
$book = Open-OfficeExcel -Path .\Finance\Purchase-Requests.xlsx -ReadOnly
try {
    Import-OfficeExcel -Document $book -Table Requests | Where-Object Status -eq 'Awaiting approval'
} finally {
    Close-OfficeExcel -Document $book
}
```

Reads the open purchase requests and closes the workbook in a finally block, so the file is released even if the filter fails.

## PARAMETERS

### -Document
Explicit-persistence workbook to close.

```yaml
Type: ExcelDocument
Parameter Sets: Default
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
