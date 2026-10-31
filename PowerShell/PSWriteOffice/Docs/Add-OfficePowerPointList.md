---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficePowerPointList
## SYNOPSIS
Adds a native list through the shared layout and continuation rules.

## SYNTAX
### Items (Default)
```powershell
Add-OfficePowerPointList [-Items] <string[]> [-Numbered] [-Level <int>] [-PassThru] [<CommonParameters>]
```

### Content
```powershell
Add-OfficePowerPointList [-Content] <scriptblock> [-Numbered] [-Level <int>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a native list through the shared layout and continuation rules.

## EXAMPLES

### EXAMPLE 1
```powershell
$list = PowerPointList 'Availability', 'Capacity' -PassThru
$list.Paragraphs[1].Text = 'Capacity planning'
```


## PARAMETERS

### -Content
Gets or sets a block of plain or rich PowerPointListItem commands.

```yaml
Type: ScriptBlock
Parameter Sets: Content
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Items
Gets or sets the list items in display order.

```yaml
Type: String[]
Parameter Sets: Items
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Level
Gets or sets the nesting level, one or two.

```yaml
Type: Int32
Parameter Sets: Items, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Numbered
Gets or sets whether the list is numbered.

```yaml
Type: SwitchParameter
Parameter Sets: Items, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Gets or sets whether to emit the native frame backing the first physical list segment.

```yaml
Type: SwitchParameter
Parameter Sets: Items, Content
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

- `None`

## OUTPUTS

- `OfficeIMO.PowerPoint.PowerPointTextBox`

## RELATED LINKS

- None
