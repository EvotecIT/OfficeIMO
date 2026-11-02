---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficePowerPointTableRow
## SYNOPSIS
Appends or inserts a row in an existing PowerPoint table.

## SYNTAX
### Table (Default)
```powershell
Add-OfficePowerPointTableRow [-InputObject] <Object> [[-Value] <Object>] [-TemplateRow <Int32>] [-Position <Int32>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Accepts a PowerPointTable or a PowerPointShapeInfo record whose shape is a
table. Values can be scalars, arrays, dictionaries, or PowerShell objects; arrays and object
properties are expanded across table cells. The new row is cloned from an existing template row so
table formatting, borders, and style choices are preserved.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Get-OfficePowerPointShape -Document $ppt -Text 'Metric' -Kind Table |
                Add-OfficePowerPointTableRow -Value 'Latency', 'Ready'
```

Accepts a PowerPoint table or table shape metadata and writes the supplied values into the new row.

### EXAMPLE 2
```powershell
PS> $shape = Get-OfficePowerPointShape -Document $ppt -Text 'Metric' -Kind Table | Select-Object -First 1
$shape | Add-OfficePowerPointTableRow -Position 2 -TemplateRow 2 -Value ([ordered]@{
    Metric = 'Documentation'
    State  = 'Ready'
})
```

Resolves the table from shape metadata, inserts a formatted row at position 2, and maps values across cells.

## PARAMETERS

### -InputObject
PowerPoint table or table shape info returned by Get-OfficePowerPointShape.

```yaml
Type: Object
Parameter Sets: Table
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
Parameter Sets: Table
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Position
Optional one-based position where the row should be inserted. Defaults to appending at the end.

```yaml
Type: Int32
Parameter Sets: Table
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TemplateRow
Optional one-based template row to clone. Defaults to the last existing row.

```yaml
Type: Int32
Parameter Sets: Table
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
Parameter Sets: Table
Aliases: None
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
