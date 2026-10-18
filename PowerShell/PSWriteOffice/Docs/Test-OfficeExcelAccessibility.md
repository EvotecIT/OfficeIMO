---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Test-OfficeExcelAccessibility
## SYNOPSIS
Checks workbook accessibility and compliance signals.

## SYNTAX
### Path (Default)
```powershell
Test-OfficeExcelAccessibility [-Path] <string> [-Quiet] [<CommonParameters>]
```

### Document
```powershell
Test-OfficeExcelAccessibility -Document <ExcelDocument> [-Quiet] [<CommonParameters>]
```

## DESCRIPTION
Checks workbook accessibility and compliance signals.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $accessibility = Test-OfficeExcelAccessibility -Path .\Dashboard.xlsx
$accessibility.Findings |
    Sort-Object Severity,Category,SheetName,Address |
    Format-Table Severity,Category,SheetName,Address,Message
```

Reports OfficeIMO accessibility findings such as missing image alt text, hidden sheets, merged ranges, and tables without header rows.

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

### -Quiet
Return only a Boolean pass/fail value.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Document
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
