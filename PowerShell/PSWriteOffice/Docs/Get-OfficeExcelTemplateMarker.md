---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeExcelTemplateMarker
## SYNOPSIS
Lists Excel template markers such as {{Name}} and optionally shows whether supplied values bind to them.

## SYNTAX
### Context (Default)
```powershell
Get-OfficeExcelTemplateMarker [-Sheet <string>] [-SheetIndex <Int32>] [-Value <hashtable>] [-MissingOnly] [<CommonParameters>]
```

### Path
```powershell
Get-OfficeExcelTemplateMarker [-Path] <string> [-Sheet <string>] [-SheetIndex <Int32>] [-Value <hashtable>] [-MissingOnly] [<CommonParameters>]
```

### Document
```powershell
Get-OfficeExcelTemplateMarker -Document <ExcelDocument> [-Sheet <string>] [-SheetIndex <Int32>] [-Value <hashtable>] [-MissingOnly] [<CommonParameters>]
```

## DESCRIPTION
Lists Excel template markers such as {{Name}} and optionally shows whether supplied values bind to them.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Get-OfficeExcelTemplateMarker -Path .\Invoice.xlsx -Sheet Invoice -Value @{ Number = 'INV-001'; Total = 123.45 }
```

Returns one object per marker with address, format, and binding metadata.

## PARAMETERS

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

### -MissingOnly
Only returns markers that are not supplied by -Value.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
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

### -Sheet
Worksheet name to inspect. Defaults to the current DSL sheet or all workbook sheets.

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

### -SheetIndex
Worksheet index (0-based) to inspect. Defaults to the current DSL sheet or all workbook sheets.

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

### -Value
Optional marker values used to report which markers are bound and which are still missing.

```yaml
Type: Hashtable
Parameter Sets: Context, Path, Document
Aliases: Values
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
