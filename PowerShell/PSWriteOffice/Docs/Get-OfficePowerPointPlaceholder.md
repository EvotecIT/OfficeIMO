---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficePowerPointPlaceholder
## SYNOPSIS
Gets native slide placeholders or metadata from the selected slide's template layout.

## SYNTAX
### Slide (Default)
```powershell
Get-OfficePowerPointPlaceholder [-Slide <PowerPointSlideSelector>] [-Document <PowerPointDocument>] [-Type <PowerPointPlaceholderType>] [-Index <UInt32>] [<CommonParameters>]
```

### Layout
```powershell
Get-OfficePowerPointPlaceholder -FromLayout [-Slide <PowerPointSlideSelector>] [-Document <PowerPointDocument>] [-Type <PowerPointPlaceholderType>] [-Index <UInt32>] [<CommonParameters>]
```

## DESCRIPTION
Gets native slide placeholders or metadata from the selected slide's template layout.

## EXAMPLES

### EXAMPLE 1
```powershell
Get-OfficePowerPointPlaceholder -Document $deck -Slide 1 -Type Title
```


### EXAMPLE 2
```powershell
$slide | Get-OfficePowerPointPlaceholder -FromLayout
```


## PARAMETERS

### -Document
Gets or sets the document for numeric selection or explicit membership checks.

```yaml
Type: PowerPointDocument
Parameter Sets: Slide, Layout
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FromLayout
Gets or sets whether to return layout metadata instead of live text boxes.

```yaml
Type: SwitchParameter
Parameter Sets: Layout
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Index
Gets or sets the native placeholder index to include.

```yaml
Type: UInt32
Parameter Sets: Slide, Layout
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Slide
Gets or sets a native slide or one-based position; omission uses the current slide block.

```yaml
Type: PowerPointSlideSelector
Parameter Sets: Slide, Layout
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Type
Gets or sets the placeholder type to include.

```yaml
Type: PowerPointPlaceholderType
Parameter Sets: Slide, Layout
Aliases: None
Possible values: Title, Body, CenteredTitle, SubTitle, DateAndTime, SlideNumber, Footer, Header, Object, Chart, Table, ClipArt, Diagram, Media, SlideImage, Picture

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `PSWriteOffice.Foundation.PowerPointSlideSelector`: A live slide or a one-based slide position in a document.

## OUTPUTS

- `OfficeIMO.PowerPoint.PowerPointTextBox`
- `OfficeIMO.PowerPoint.PowerPointLayoutPlaceholderInfo`

## RELATED LINKS

- None
