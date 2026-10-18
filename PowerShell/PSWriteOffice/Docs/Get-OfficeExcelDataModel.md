---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeExcelDataModel
## SYNOPSIS
Inspects workbook data model, query, connection, and external-link package parts.

## SYNTAX
### Path (Default)
```powershell
Get-OfficeExcelDataModel [-Path] <string> [<CommonParameters>]
```

### Document
```powershell
Get-OfficeExcelDataModel -Document <ExcelDocument> [<CommonParameters>]
```

## DESCRIPTION
Inspects workbook data model, query, connection, and external-link package parts.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $model = Get-OfficeExcelDataModel -Path .\WorkbookWithQueries.xlsx
if ($model.HasDataModelOrQueries) {
    $model.Details | ForEach-Object { "Preserved package part: $_" }
}
```

Separates preserved connection/query/data-model parts from executable refresh behavior so automation can decide whether Excel refresh is required.

## PARAMETERS

### -Document
Workbook document.

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
Workbook path.

```yaml
Type: String
Parameter Sets: Path
Aliases: InputPath, FilePath
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

- `System.Management.Automation.PSObject`

## RELATED LINKS

- None
