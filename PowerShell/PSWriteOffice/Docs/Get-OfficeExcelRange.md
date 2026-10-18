---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeExcelRange
## SYNOPSIS
Reads an explicit A1 range from an Excel workbook.

## SYNTAX
### Path (Default)
```powershell
Get-OfficeExcelRange [-Path] <string> -Range <string> [-Sheet <string>] [-SheetIndex <Int32>] [-HeadersInFirstRow <bool>] [-NumericAsDecimal] [-AsHashtable] [-AsDataTable] [<CommonParameters>]
```

### Uri
```powershell
Get-OfficeExcelRange [-Uri] <uri> -Range <string> [-AllowHttp] [-Sheet <string>] [-SheetIndex <Int32>] [-HeadersInFirstRow <bool>] [-NumericAsDecimal] [-AsHashtable] [-AsDataTable] [<CommonParameters>]
```

### Document
```powershell
Get-OfficeExcelRange -Document <ExcelDocument> -Range <string> [-Sheet <string>] [-SheetIndex <Int32>] [-HeadersInFirstRow <bool>] [-NumericAsDecimal] [-AsHashtable] [-AsDataTable] [<CommonParameters>]
```

## DESCRIPTION
Returns rows as PSCustomObjects by default, with optional hashtable or DataTable output for scripting and interoperability.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $rows = Get-OfficeExcelRange -Path .\report.xlsx -Sheet Data -Range A1:C10 -NumericAsDecimal
$rows |
    Group-Object -Property Region |
    Select-Object Name, Count
```

Uses the first row as headers, returns row objects, and keeps numeric values decimal-friendly.

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

### -AsDataTable
Emit the raw DataTable instead of row objects.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Uri, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AsHashtable
Emit rows as hashtables instead of PSCustomObjects.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Uri, Document
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

### -HeadersInFirstRow
Use the first row as column headers.

```yaml
Type: Boolean
Parameter Sets: Path, Uri, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NumericAsDecimal
Prefer decimals instead of doubles for numeric values.

```yaml
Type: SwitchParameter
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

### -Range
A1 range to read.

```yaml
Type: String
Parameter Sets: Path, Uri, Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Sheet
Worksheet name to read; defaults to the first sheet.

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
Zero-based worksheet index to read; defaults to the first sheet.

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
Remote workbook URI to read.

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
- `System.Collections.Hashtable`
- `System.Data.DataTable`

## RELATED LINKS

- None
