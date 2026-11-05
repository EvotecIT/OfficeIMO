---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficeExcelTemplateMarker
## SYNOPSIS
Lists template markers and, with -Value, which are bound.

## SYNTAX
### Document (Default)
```powershell
Get-OfficeExcelTemplateMarker [-Sheet <OfficeSheetSelector[]>] [-Value <hashtable>] [-MissingOnly] [-Document <ExcelDocument>] [<CommonParameters>]
```

### Path
```powershell
Get-OfficeExcelTemplateMarker [-Path] <string[]> [-Sheet <OfficeSheetSelector[]>] [-Value <hashtable>] [-MissingOnly] [<CommonParameters>]
```

### LiteralPath
```powershell
Get-OfficeExcelTemplateMarker -LiteralPath <string[]> [-Sheet <OfficeSheetSelector[]>] [-Value <hashtable>] [-MissingOnly] [<CommonParameters>]
```

## DESCRIPTION
Lists template markers and, with -Value, which are bound.

## EXAMPLES

### EXAMPLE 1
```powershell
Get-OfficeExcelTemplateMarker -Path .\Templates\Service-Report.xltx
```

Shows every {{Marker}} in the template with its sheet and cell.

### EXAMPLE 2
```powershell
$values = @{ Client = 'Northwind Traders'; Month = 'October 2026' }
Get-OfficeExcelTemplateMarker -Path .\Templates\Service-Report.xltx -Value $values -MissingOnly |
    Select-Object SheetName, CellReference, Name
```

Compares the template with the values a script is about to supply and lists only the markers that would stay unfilled.

## PARAMETERS

### -Document
Caller-owned workbook; omission uses the enclosing Excel content block.

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
Literal workbook paths, including FileInfo.FullName and provider PSPath.

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

### -MissingOnly
Only returns markers that are not supplied by -Value.

```yaml
Type: SwitchParameter
Parameter Sets: Document, Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Workbook paths resolved with filesystem wildcard support.

```yaml
Type: String[]
Parameter Sets: Path
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Sheet
Worksheet name, one-based position or native worksheet object. Multi-sheet commands also accept wildcard names.

```yaml
Type: OfficeSheetSelector[]
Parameter Sets: Document, Path, LiteralPath
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
Type: Hashtable
Parameter Sets: Document, Path, LiteralPath
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

- `System.String[]`
- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `None`

## RELATED LINKS

- None
