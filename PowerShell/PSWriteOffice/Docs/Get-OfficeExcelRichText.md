---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeExcelRichText
## SYNOPSIS
Gets mixed-format rich text runs from an Excel cell.

## SYNTAX
### Context (Default)
```powershell
Get-OfficeExcelRichText [-Sheet <string>] [-SheetIndex <Int32>] [-Row <Int32>] [-Column <Int32>] [-Address <string>] [<CommonParameters>]
```

### Path
```powershell
Get-OfficeExcelRichText [-Path] <string> [-Sheet <string>] [-SheetIndex <Int32>] [-Row <Int32>] [-Column <Int32>] [-Address <string>] [<CommonParameters>]
```

### Document
```powershell
Get-OfficeExcelRichText -Document <ExcelDocument> [-Sheet <string>] [-SheetIndex <Int32>] [-Row <Int32>] [-Column <Int32>] [-Address <string>] [<CommonParameters>]
```

## DESCRIPTION
Gets mixed-format rich text runs from an Excel cell.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $runs = Get-OfficeExcelRichText -Path .\Report.xlsx -Sheet Summary -Address A1
$runs |
    Format-Table Index, Text, Bold, Italic, Color
```

Returns one object per inline rich text run with text and style properties.

## PARAMETERS

### -Address
A1-style cell address.

```yaml
Type: String
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Column
1-based column index.

```yaml
Type: Int32
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Workbook to inspect outside the DSL context.

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
Aliases: InputPath, FilePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Row
1-based row index.

```yaml
Type: Int32
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Sheet
Worksheet name. Defaults to the current sheet inside an ExcelSheet block.

```yaml
Type: String
Parameter Sets: Context, Path, Document
Aliases: WorksheetName
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SheetIndex
Worksheet index when using a workbook object or path.

```yaml
Type: Int32
Parameter Sets: Context, Path, Document
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

- `System.Management.Automation.PSObject`

## RELATED LINKS

- None
