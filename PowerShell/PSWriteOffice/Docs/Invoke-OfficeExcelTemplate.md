---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Invoke-OfficeExcelTemplate
## SYNOPSIS
Fills {{Marker}} placeholders, repeats template rows or sheets, and keeps or removes optional rows.

## SYNTAX
### Values (Default)
```powershell
Invoke-OfficeExcelTemplate [-Value] <hashtable> [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector[]>] [-Culture <cultureinfo>] [-MissingValueBehavior <ExcelTemplateMissingValueBehavior>] [-ThrowOnMissing] [-PassThru] [<CommonParameters>]
```

### Rows
```powershell
Invoke-OfficeExcelTemplate [-InputObject] <Object> -TemplateRow <Int32> [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector[]>] [-Culture <cultureinfo>] [-MissingValueBehavior <ExcelTemplateMissingValueBehavior>] [-ThrowOnMissing] [-PassThru] [<CommonParameters>]
```

### OptionalRows
```powershell
Invoke-OfficeExcelTemplate -OptionalRow <Int32> [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector[]>] [-Value <hashtable>] [-Culture <cultureinfo>] [-MissingValueBehavior <ExcelTemplateMissingValueBehavior>] [-ThrowOnMissing] [-PassThru] [-RowCount <int>] [-Remove] [<CommonParameters>]
```

### Sheets
```powershell
Invoke-OfficeExcelTemplate [-InputObject] <Object> -TemplateSheet <OfficeSheetSelector> [-Document <ExcelDocument>] [-Culture <cultureinfo>] [-MissingValueBehavior <ExcelTemplateMissingValueBehavior>] [-ThrowOnMissing] [-PassThru] [-SheetNameProperty <string>] [<CommonParameters>]
```

## DESCRIPTION
Fills {{Marker}} placeholders, repeats template rows or sheets, and keeps or removes optional rows.

## EXAMPLES

### EXAMPLE 1
```powershell
New-OfficeExcel -Path .\Reports\Service-Report-October.xlsx -Template .\Templates\Service-Report.xltx {
    ExcelTemplate @{ Client = 'Northwind Traders'; Month = 'October 2026'; Availability = 0.9994 }
}
```

Creates a report from a template and replaces its {{Client}}, {{Month}} and {{Availability}} markers.

### EXAMPLE 2
```powershell
$incidents = Import-OfficeCsv -Path .\Data\Incidents-October.csv
$values = @{
    Client       = 'Northwind Traders'
    Month        = 'October 2026'
    Availability = 0.9994
}
New-OfficeExcel -Path .\Reports\Service-Report-October.xlsx -Template .\Templates\Service-Report.xltx {
    ExcelTemplate $values -Culture en-GB
    ExcelSheet 'Report' {
        ExcelTemplate -TemplateRow 12 -InputObject $incidents
        ExcelTemplate -OptionalRow 18 -RowCount 3 -Remove
    }
}
```

Fills the header markers, repeats row 12 once per incident and removes the escalation block because nothing was escalated this month.

### EXAMPLE 3
```powershell
$sites = Import-OfficeCsv -Path .\Facilities\Sites.csv
Open-OfficeExcel -Path .\Facilities\Inspections-2026.xlsx {
    ExcelTemplate -TemplateSheet Site -InputObject $sites -SheetNameProperty Name -ThrowOnMissing
}
```

Copies the Site template sheet once per site, names each copy after the site and fills its markers from the site record.

## PARAMETERS

### -Culture
Culture used for value conversion and number-format presets.

```yaml
Type: CultureInfo
Parameter Sets: Values, Rows, OptionalRows, Sheets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Caller-owned workbook; omission uses the enclosing Excel content block.

```yaml
Type: ExcelDocument
Parameter Sets: Values, Rows, OptionalRows, Sheets
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
Parameter Sets: Rows, Sheets
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -MissingValueBehavior
Behavior used when a marker is not supplied by -Value.

```yaml
Type: ExcelTemplateMissingValueBehavior
Parameter Sets: Values, Rows, OptionalRows, Sheets
Aliases: None
Possible values: PreserveMarker, EmptyString, Throw

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OptionalRow
Template row blocks whose visibility or inclusion is controlled by the supplied values.

```yaml
Type: Int32
Parameter Sets: OptionalRows
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emits the native result after the operation.

```yaml
Type: SwitchParameter
Parameter Sets: Values, Rows, OptionalRows, Sheets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Remove
Removes the optional row block instead of keeping and binding it.

```yaml
Type: SwitchParameter
Parameter Sets: OptionalRows
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RowCount
Number of rows in the optional block.

```yaml
Type: Int32
Parameter Sets: OptionalRows
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
Type: OfficeSheetSelector[]
Parameter Sets: Values, Rows, OptionalRows
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SheetNameProperty
Input property used as the generated worksheet name.

```yaml
Type: String
Parameter Sets: Sheets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TemplateRow
1-based row number that contains template markers to repeat.

```yaml
Type: Int32
Parameter Sets: Rows
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TemplateSheet
Template worksheet name. Defaults to the current sheet inside an ExcelSheet block or the first sheet for path/document use.

```yaml
Type: OfficeSheetSelector
Parameter Sets: Sheets
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ThrowOnMissing
Throws when a marker is not supplied by -Value.

```yaml
Type: SwitchParameter
Parameter Sets: Values, Rows, OptionalRows, Sheets
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
Parameter Sets: Values, OptionalRows
Aliases: None
Possible values:

Required: False
Position: 0
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
