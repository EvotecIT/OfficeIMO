---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Edit-OfficeExcelRow
## SYNOPSIS
Runs a script block against each editable row of a range and writes the changes back.

## SYNTAX
### Document (Default)
```powershell
Edit-OfficeExcelRow [-ScriptBlock] <scriptblock> [-Sheet <OfficeSheetSelector>] [-Range <string>] [-NumericAsDecimal] [-PassThru] [-Document <ExcelDocument>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Path
```powershell
Edit-OfficeExcelRow [-Path] <string[]> [-ScriptBlock] <scriptblock> [-Sheet <OfficeSheetSelector>] [-Range <string>] [-NumericAsDecimal] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPath
```powershell
Edit-OfficeExcelRow [-ScriptBlock] <scriptblock> -LiteralPath <string[]> [-Sheet <OfficeSheetSelector>] [-Range <string>] [-NumericAsDecimal] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Runs a script block against each editable row of a range and writes the changes back.

## EXAMPLES

### EXAMPLE 1
```powershell
Edit-OfficeExcelRow -Path .\Reports\Releases.xlsx -Sheet Releases -ScriptBlock {
    param($row)
    if ($row['Status'].Value -eq 'Draft') { $row.Set('Status', 'Ready') }
}
```

Marks every draft row on the Releases sheet as ready and saves the workbook.

### EXAMPLE 2
```powershell
$book = Open-OfficeExcel -Path .\Contracts\Licences.xlsx
Edit-OfficeExcelRow -Document $book -Sheet Licences -Range A1:F400 -ScriptBlock {
    param($row)
    if ($row['State'].Value -eq 'Active') {
        $row.Set('RenewalDate', $row.Get[datetime]('StartDate').AddYears(1))
    }
}
Save-OfficeExcel -Document $book -PassThru | Close-OfficeExcel
```

Sets the renewal date of each active licence to one year after its start date in a live workbook, which the script then saves.

### EXAMPLE 3
```powershell
$tickets = Import-OfficeCsv -Path .\Data\Tickets.csv
New-OfficeExcel -Path .\Reports\Tickets.xlsx {
    ExcelSheet 'Tickets' {
        ExcelTable $tickets -Name Tickets
    }
    Edit-OfficeExcelRow -Sheet Tickets -ScriptBlock {
        param($row)
        if (-not $row['Owner'].Value) { $row.Set('Owner', 'Unassigned') }
    }
}
```

Writes imported rows and then fixes up their status inside the same block, without a -Document variable.

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

### -NumericAsDecimal
Prefer decimals instead of doubles for numeric values.

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

### -PassThru
Emits the native result after the operation.

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

### -Range
A1 cell range on the selected worksheet.

```yaml
Type: String
Parameter Sets: Document, Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ScriptBlock
Script block to run once per editable row. The row is passed as the first argument.

```yaml
Type: ScriptBlock
Parameter Sets: Document, Path, LiteralPath
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Sheet
Worksheet name, one-based position or native worksheet object. Multi-sheet commands also accept wildcard names.

```yaml
Type: OfficeSheetSelector
Parameter Sets: Document, Path, LiteralPath
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

- `System.String[]`
- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `None`

## RELATED LINKS

- None
