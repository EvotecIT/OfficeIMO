---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficeExcelSheetView
## SYNOPSIS
Gets frozen panes, gridlines, direction, zoom and view mode for worksheets.

## SYNTAX
### Document (Default)
```powershell
Get-OfficeExcelSheetView [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector[]>] [<CommonParameters>]
```

### Path
```powershell
Get-OfficeExcelSheetView [-Path] <string[]> [-Sheet <OfficeSheetSelector[]>] [<CommonParameters>]
```

### LiteralPath
```powershell
Get-OfficeExcelSheetView -LiteralPath <string[]> [-Sheet <OfficeSheetSelector[]>] [<CommonParameters>]
```

## DESCRIPTION
Gets frozen panes, gridlines, direction, zoom and view mode for worksheets.

## EXAMPLES

### EXAMPLE 1
```powershell
Get-ChildItem .\Reports -Filter *.xlsx | Get-OfficeExcelSheetView -Sheet Summary | Format-Table
```


## PARAMETERS

### -Document
Gets or sets a live workbook; omission uses the current worksheet or workbook scope.

```yaml
Type: ExcelDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -LiteralPath
Gets or sets literal workbook paths, including FileInfo.FullName and provider PSPath.

```yaml
Type: String[]
Parameter Sets: LiteralPath
Aliases: PSPath, FullName
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -Path
Gets or sets workbook paths with filesystem wildcard expansion.

```yaml
Type: String[]
Parameter Sets: Path
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -Sheet
Gets or sets worksheet names or wildcard patterns, one-based positions or native worksheet objects.

```yaml
Type: OfficeSheetSelector[]
Parameter Sets: Document, Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.String[]`
- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `PSWriteOffice.Models.Excel.ExcelSheetViewInfo`: A worksheet view snapshot with its workbook path and sheet name.

## RELATED LINKS

- None
