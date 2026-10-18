---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeExcelWriteReservation
## SYNOPSIS
Gets workbook write-reservation metadata.

## SYNTAX
### Context (Default)
```powershell
Get-OfficeExcelWriteReservation [<CommonParameters>]
```

### Path
```powershell
Get-OfficeExcelWriteReservation [-Path] <string> [<CommonParameters>]
```

### Document
```powershell
Get-OfficeExcelWriteReservation -Document <ExcelDocument> [<CommonParameters>]
```

## DESCRIPTION
Gets workbook write-reservation metadata.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Get-OfficeExcelWriteReservation -Path .\Report.xlsx |
                Format-List Exists, ReadOnlyRecommended, UserName, HasPasswordHash
```

Reports Excel file-sharing/write-reservation metadata separately from workbook protection and package encryption.

## PARAMETERS

### -Document
Open workbook document to inspect.

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

### -Path
Workbook path to inspect.

```yaml
Type: String
Parameter Sets: Path
Aliases: FilePath, InputPath, FullName
Possible values:

Required: True
Position: 0
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
