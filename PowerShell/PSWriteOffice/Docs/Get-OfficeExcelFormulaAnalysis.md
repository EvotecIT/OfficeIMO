---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeExcelFormulaAnalysis
## SYNOPSIS
Gets workbook formula references, functions, volatile formulas, and external links.

## SYNTAX
### Path (Default)
```powershell
Get-OfficeExcelFormulaAnalysis [-Path] <string> [-IncludeFormulas] [<CommonParameters>]
```

### Document
```powershell
Get-OfficeExcelFormulaAnalysis -Document <ExcelDocument> [-IncludeFormulas] [<CommonParameters>]
```

## DESCRIPTION
Gets workbook formula references, functions, volatile formulas, and external links.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $analysis = Get-OfficeExcelFormulaAnalysis -Path .\Model.xlsx -IncludeFormulas
$analysis.Formulas |
    Where-Object { $_.IsVolatile -or $_.HasExternalReference } |
    Format-Table SheetName,Address,Formula,IsVolatile,HasExternalReference
```

Uses OfficeIMO formula analysis so scripts can review volatile functions and external workbook references without parsing package XML themselves.

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

### -IncludeFormulas
Include per-cell formula details.

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
