---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
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

### Content
```powershell
Set-OfficePowerPointTableCell [-InputObject] <Object> -Row <int> -Column <int> -Content <scriptblock> [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Accepts a PowerPointTable or a PowerPointShapeInfo record whose shape is a
table. Row and column coordinates are one-based. Use
this after Get-OfficePowerPointShape -Kind Table when a script needs to update a specific
cell inside a deck that already exists.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Get-OfficePowerPointShape -Document $ppt -Text 'Metric' -Kind Table |
                Set-OfficePowerPointTableCell -Row 2 -Column 2 -Text 'Ready'
```

Accepts a PowerPoint table or table shape metadata and updates a one-based table cell.

### EXAMPLE 2
```powershell
PS> $table = Get-OfficePowerPointShape -Document $ppt -Text 'Risk' -Kind Table | Select-Object -First 1
$table | Set-OfficePowerPointTableCell -Row 2 -Column 2 -Text 'Mitigating'
```

Uses the table found by text content and updates the second row, second column.

## PARAMETERS

### -Column
One-based column number.

```yaml
Type: Int32
Parameter Sets: Text, Content
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Content
Replacement rich paragraphs and runs, using PowerPointParagraph and PowerPointText.

```yaml
Type: ScriptBlock
Parameter Sets: Content
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -InputObject
PowerPoint table or table shape info returned by Get-OfficePowerPointShape.

```yaml
Type: Object
Parameter Sets: Text, Content
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
Parameter Sets: Text, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Row
One-based row number.

```yaml
Type: Int32
Parameter Sets: Text, Content
Aliases: None
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
