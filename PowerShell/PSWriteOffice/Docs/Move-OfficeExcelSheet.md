---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Move-OfficeExcelSheet
## SYNOPSIS
Moves a worksheet to a new workbook position.

## SYNTAX
### Context (Default)
```powershell
Move-OfficeExcelSheet -Index <int> [-Sheet <string>] [-SheetIndex <Int32>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Path
```powershell
Move-OfficeExcelSheet [-Path] <string> -Index <int> [-Sheet <string>] [-SheetIndex <Int32>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
Move-OfficeExcelSheet -Document <ExcelDocument> -Index <int> [-Sheet <string>] [-SheetIndex <Int32>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Moves a worksheet to a new workbook position.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $proof = @(
    Move-OfficeExcelSheet -Path .\Report.xlsx -Sheet Summary -Index 0
    Get-OfficeExcelSummary -Path .\Report.xlsx -IncludeSheets |
        Select-Object -ExpandProperty Sheets |
        Select-Object -First 3 -Property Index, Name
)
$proof
```

Moves Summary to the first worksheet tab and reads back the first sheets from workbook summary.

## PARAMETERS

### -Document
Workbook to update outside the DSL context.

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

### -Index
Zero-based destination tab index.

```yaml
Type: Int32
Parameter Sets: Context, Path, Document
Aliases: TargetIndex
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the moved worksheet.

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
Workbook path to update.

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
Worksheet name to move. Defaults to the current sheet inside an ExcelSheet block.

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
Worksheet index to move when using a workbook object or path.

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

- `None`

## RELATED LINKS

- None
