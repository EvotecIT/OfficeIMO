---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficeExcelNamedRange
## SYNOPSIS
Gets defined names from a workbook.

## SYNTAX
### Document (Default)
```powershell
Get-OfficeExcelNamedRange [-Name <string>] [-Sheet <OfficeSheetSelector>] [-Document <ExcelDocument>] [<CommonParameters>]
```

### Path
```powershell
Get-OfficeExcelNamedRange [-Path] <string[]> [-Name <string>] [-Sheet <OfficeSheetSelector>] [<CommonParameters>]
```

### LiteralPath
```powershell
Get-OfficeExcelNamedRange -LiteralPath <string[]> [-Name <string>] [-Sheet <OfficeSheetSelector>] [<CommonParameters>]
```

### Uri
```powershell
Get-OfficeExcelNamedRange [-Uri] <uri> [-Name <string>] [-Sheet <OfficeSheetSelector>] [-AllowHttp] [<CommonParameters>]
```

## DESCRIPTION
Gets defined names from a workbook.

## EXAMPLES

### EXAMPLE 1
```powershell
Get-OfficeExcelNamedRange -Path .\Finance\Pricing-Model.xlsx
```

Shows every defined name and what it refers to.

### EXAMPLE 2
```powershell
Get-OfficeExcelNamedRange -Path .\Finance\Budget-2027.xlsx |
                Where-Object Reference -like '*#REF!*' |
                Select-Object Name, SheetName, Reference
```

Lists the defined names that refer to #REF! after sheets were deleted, so they can be repaired or removed.

## PARAMETERS

### -AllowHttp
Allow HTTP workbook downloads in addition to HTTPS.

```yaml
Type: SwitchParameter
Parameter Sets: Uri
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

### -Name
Optional named range to retrieve.

```yaml
Type: String
Parameter Sets: Document, Path, LiteralPath, Uri
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
Type: OfficeSheetSelector
Parameter Sets: Document, Path, LiteralPath, Uri
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Uri
Remote workbook URI to inspect.

```yaml
Type: Uri
Parameter Sets: Uri
Aliases: None
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

- `System.String[]`
- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `None`

## RELATED LINKS

- None
