---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficePowerPointTableRow
## SYNOPSIS
Appends or inserts a row in an existing PowerPoint table.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficePowerPointTableRow [-InputObject] <Object> [[-Value] <Object>] [-TemplateRowIndex <Int32>] [-Index <Int32>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Accepts a PowerPointTable or a PowerPointShapeInfo record whose shape is a
table. Values can be scalars, arrays, dictionaries, or PowerShell objects; arrays and object
properties are expanded across table cells. The new row is cloned from an existing template row so
table formatting, borders, and style choices are preserved.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Find-OfficePowerPointShape -Presentation $ppt -Text 'Metric' -Kind Table |
                Add-OfficePowerPointTableRow -Values 'Latency', 'Ready'
```

Accepts a PowerPoint table or table shape metadata and writes the supplied values into the new row.

### EXAMPLE 2
```powershell
PS> $shape = Find-OfficePowerPointShape -Presentation $ppt -Text 'Metric' -Kind Table | Select-Object -First 1
$shape | Add-OfficePowerPointTableRow -Index 1 -TemplateRowIndex 1 -Values ([ordered]@{
    Metric = 'Documentation'
    State  = 'Ready'
})
```

Resolves the table from shape metadata, inserts a formatted row at index 1, and maps values across cells.

## PARAMETERS

### -Index
Optional zero-based index where the row should be inserted. Defaults to appending at the end.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -InputObject
PowerPoint table or table shape info returned by Find-OfficePowerPointShape or Get-OfficePowerPointShape.

```yaml
Type: Object
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -PassThru
Emit the created row for additional OfficeIMO-level edits.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TemplateRowIndex
Optional zero-based template row index to clone. Defaults to the last existing row.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Value
Values to write into the new row. Arrays, dictionaries, and objects are expanded across cells.

```yaml
Type: Object
Parameter Sets: __AllParameterSets
Aliases: Data, Values
Possible values:

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.Object`

## OUTPUTS

- `OfficeIMO.PowerPoint.PowerPointTableRow`

## RELATED LINKS

- None
