---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficePowerPointTableCell
## SYNOPSIS
Sets text in an existing PowerPoint table cell.

## SYNTAX
### Text (Default)
```powershell
Set-OfficePowerPointTableCell [-InputObject] <Object> -Row <int> -Column <int> -Text <string> [-PassThru] [<CommonParameters>]
```

### Run
```powershell
Set-OfficePowerPointTableCell [-InputObject] <Object> -Row <int> -Column <int> -Run <Object[]> [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Accepts a PowerPointTable or a PowerPointShapeInfo record whose shape is a
table. Row and column coordinates are zero-based, matching the OfficeIMO PowerPoint table API. Use
this after Find-OfficePowerPointShape -Kind Table when a script needs to update a specific
cell inside a deck that already exists.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Find-OfficePowerPointShape -Presentation $ppt -Text 'Metric' -Kind Table |
                Set-OfficePowerPointTableCell -Row 1 -Column 1 -Text 'Ready'
```

Accepts a PowerPoint table or table shape metadata and updates a zero-based table cell.

### EXAMPLE 2
```powershell
PS> $table = Find-OfficePowerPointShape -Presentation $ppt -Text 'Risk' -Kind Table | Select-Object -First 1
$table | Set-OfficePowerPointTableCell -Row 1 -Column 1 -Text 'Mitigating'
```

Uses the table found by text content and updates the second row, second column.

## PARAMETERS

### -Column
Zero-based column index.

```yaml
Type: Int32
Parameter Sets: Text, Run
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -InputObject
PowerPoint table or table shape info returned by Find-OfficePowerPointShape or Get-OfficePowerPointShape.

```yaml
Type: Object
Parameter Sets: Text, Run
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -PassThru
Emit the updated table cell for additional OfficeIMO-level edits.

```yaml
Type: SwitchParameter
Parameter Sets: Text, Run
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Row
Zero-based row index.

```yaml
Type: Int32
Parameter Sets: Text, Run
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Run
Replacement rich text runs. Each run can be created with TextRun/PowerPointTextRun or provided as a hashtable/object.

```yaml
Type: Object[]
Parameter Sets: Run
Aliases: Runs
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Text
Replacement cell text. A null value clears the cell.

```yaml
Type: String
Parameter Sets: Text
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

- `System.Object`

## OUTPUTS

- `OfficeIMO.PowerPoint.PowerPointTableCell`

## RELATED LINKS

- None
