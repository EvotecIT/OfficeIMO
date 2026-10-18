---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeExcelStreamingContract
## SYNOPSIS
Reports large-workbook streaming and direct-writer suitability.

## SYNTAX
### Path (Default)
```powershell
Get-OfficeExcelStreamingContract [-Path] <string> [<CommonParameters>]
```

### Document
```powershell
Get-OfficeExcelStreamingContract -Document <ExcelDocument> [<CommonParameters>]
```

## DESCRIPTION
Reports large-workbook streaming and direct-writer suitability.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $contract = Get-OfficeExcelStreamingContract -Path .\LargeExport.xlsx
[pscustomobject]@{
    EstimatedCells = $contract.EstimatedCellCount
    DirectWriter   = $contract.HasDirectDataSetFastSaveState
    Recommendation = $contract.Recommendation
}
```

Reports whether the workbook is already using OfficeIMO direct tabular state and gives a size-based recommendation for future imports or exports.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `System.Management.Automation.PSObject`

## RELATED LINKS

- None
