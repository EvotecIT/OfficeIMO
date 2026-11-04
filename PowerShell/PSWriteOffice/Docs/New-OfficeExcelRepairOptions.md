---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficeExcelRepairOptions
## SYNOPSIS
Creates Excel repair settings for Repair-OfficeDocument.

## SYNTAX
### Default (Default)
```powershell
New-OfficeExcelRepairOptions [-Area <ExcelRepairArea[]>] [<CommonParameters>]
```

## DESCRIPTION
Creates Excel repair settings for Repair-OfficeDocument.

## EXAMPLES

### EXAMPLE 1
```powershell
Repair-OfficeDocument -Path .\Reports\Broken.xlsx -OutputPath .\Reports\Repaired.xlsx -Options (New-OfficeExcelRepairOptions -Area DefinedNames, Tables)
```

Repairs only the defined names and tables of a damaged workbook and writes the result to a new file.

### EXAMPLE 2
```powershell
$repair = New-OfficeExcelRepairOptions -Area SheetViews, PrintSettings
Get-ChildItem -Path \\files\Migrated -Filter *.xlsx |
    Repair-OfficeDocument -Options $repair -PassThru
```

Fixes sheet views and print settings in every workbook of a migrated share, leaving formulas and drawings untouched.

## PARAMETERS

### -Area
Native workbook areas included in the operation; omitted selection uses the native default.

```yaml
Type: ExcelRepairArea[]
Parameter Sets: Default
Aliases: None
Possible values: DefinedNames, Tables, SheetViews, PrintSettings, Protection, AutoFilters, Drawings, LinksAndComments, Calculation, WorkbookArtifacts

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `None`

## RELATED LINKS

- None
