---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Test-OfficeExcelTemplate
## SYNOPSIS
Tests whether supplied values bind every template marker.

## SYNTAX
### Document (Default)
```powershell
Test-OfficeExcelTemplate -Value <hashtable> [-Quiet] [-Document <ExcelDocument>] [<CommonParameters>]
```

### Path
```powershell
Test-OfficeExcelTemplate [-Path] <string[]> -Value <hashtable> [-Quiet] [<CommonParameters>]
```

### LiteralPath
```powershell
Test-OfficeExcelTemplate -Value <hashtable> -LiteralPath <string[]> [-Quiet] [<CommonParameters>]
```

## DESCRIPTION
Tests whether supplied values bind every template marker.

## EXAMPLES

### EXAMPLE 1
```powershell
Test-OfficeExcelTemplate -Path .\Templates\Service-Report.xltx -Value @{ Client = 'Northwind Traders'; Month = 'October 2026'; Availability = 0.9994 } -Quiet
```

Returns $true when the values bind every marker in the template.

### EXAMPLE 2
```powershell
$values = @{ Client = 'Northwind Traders'; Month = 'October 2026'; Availability = 0.9994 }
$report = Test-OfficeExcelTemplate -Path .\Templates\Service-Report.xltx -Value $values
$report.ToMarkdown() | Set-Content -Path .\Logs\Template-Binding.md
if (-not (Test-OfficeExcelTemplate -Path .\Templates\Service-Report.xltx -Value $values -Quiet)) {
    throw 'The service report template has markers this script does not fill. See Logs\Template-Binding.md.'
}
```

Writes a Markdown binding report for review and stops before generating the monthly reports if any marker is unbound.

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

### -Quiet
Returns a Boolean result instead of the native validation report.

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

### -Value
Value or values consumed by the selected operation.

```yaml
Type: Hashtable
Parameter Sets: Document, Path, LiteralPath
Aliases: None
Possible values:

Required: True
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
