---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeExcelNamedRange
## SYNOPSIS
Gets defined names (named ranges) from an Excel workbook.

## SYNTAX
### Path (Default)
```powershell
Get-OfficeExcelNamedRange [-Path] <string> [-Name <string>] [-Sheet <string>] [-SheetIndex <Int32>] [<CommonParameters>]
```

### Uri
```powershell
Get-OfficeExcelNamedRange [-Uri] <uri> [-AllowHttp] [-Name <string>] [-Sheet <string>] [-SheetIndex <Int32>] [<CommonParameters>]
```

### Document
```powershell
Get-OfficeExcelNamedRange -Document <ExcelDocument> [-Name <string>] [-Sheet <string>] [-SheetIndex <Int32>] [<CommonParameters>]
```

## DESCRIPTION
Gets defined names (named ranges) from an Excel workbook.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $ranges = Get-OfficeExcelNamedRange -Path .\report.xlsx
$ranges |
    Select-Object -Property Name, Scope, Range |
    Format-Table -AutoSize
```

Returns named ranges so workbook navigation and formulas can be reviewed.

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
Workbook to inspect.

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

### -Name
Optional named range to retrieve.

```yaml
Type: String
Parameter Sets: Path, Uri, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Path to the workbook.

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

### -Sheet
Optional sheet name for sheet-scoped names.

```yaml
Type: String
Parameter Sets: Path, Uri, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SheetIndex
Optional sheet index (0-based) for sheet-scoped names.

```yaml
Type: Int32
Parameter Sets: Path, Uri, Document
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
Aliases: Url
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
