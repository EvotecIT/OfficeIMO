---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficeExcelTableRow
## SYNOPSIS
Appends rows to an existing Excel table.

## SYNTAX
### Default (Default)
```powershell
Add-OfficeExcelTableRow [-InputObject] <Object> -Table <Object> [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Appends rows to an existing Excel table.

## EXAMPLES

### EXAMPLE 1
```powershell
$servers = @(
    [pscustomobject]@{ Name = 'NW-APP-07'; Role = 'Application'; Site = 'Warsaw' }
    [pscustomobject]@{ Name = 'NW-SQL-03'; Role = 'Database'; Site = 'Frankfurt' }
)
Open-OfficeExcel -Path .\Reports\Inventory.xlsx {
    Add-OfficeExcelTableRow -Sheet Servers -Table Servers -InputObject $servers
}
```

Adds two new servers to the Servers table; the table range grows to include them.

### EXAMPLE 2
```powershell
$book = Open-OfficeExcel -Path .\Reports\Deployments.xlsx
foreach ($target in 'NW-APP-01', 'NW-APP-02', 'NW-APP-03') {
    $result = Invoke-Command -ComputerName $target { & C:\Deploy\Install.ps1 -Release '2026.10.2' }
    [pscustomobject]@{ Server = $target; Release = '2026.10.2'; Finished = Get-Date; Status = $result.Status } |
        Add-OfficeExcelTableRow -Document $book -Sheet Log -Table Deployments
}
Save-OfficeExcel -Document $book -PassThru | Close-OfficeExcel
```

Appends one row per completed deployment to the Deployments table of a workbook kept open for the whole run.

## PARAMETERS

### -Document
Caller-owned workbook; omission uses the enclosing Excel content block.

```yaml
Type: ExcelDocument
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -InputObject
Objects, dictionaries or native tabular data supplied directly or through the pipeline.

```yaml
Type: Object
Parameter Sets: Default
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -PassThru
Emits the native result after the operation.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Sheet
Worksheet name, one-based position or native worksheet object. Multi-sheet commands also accept wildcard names.

```yaml
Type: OfficeSheetSelector
Parameter Sets: Default
Aliases: WorksheetName
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Table
Native Excel table used to resolve the source range or header.

```yaml
Type: Object
Parameter Sets: Default
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

- `OfficeIMO.Excel.ExcelDocument`
- `System.Object`

## OUTPUTS

- `None`

## RELATED LINKS

- None
