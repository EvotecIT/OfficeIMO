---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Set-OfficePowerPointPlaceholder
## SYNOPSIS
Replaces a native slide placeholder with plain text or rich paragraph content.

## SYNTAX
### Text (Default)
```powershell
Set-OfficePowerPointPlaceholder [-Text] <string> -Type <PowerPointPlaceholderType> [-Slide <PowerPointSlideSelector>] [-Index <UInt32>] [-IgnoreMissing] [-PassThru] [<CommonParameters>]
```

### Content
```powershell
Set-OfficePowerPointPlaceholder [-Content] <scriptblock> -Type <PowerPointPlaceholderType> [-Slide <PowerPointSlideSelector>] [-Index <UInt32>] [-IgnoreMissing] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Edits the existing text box, preserving its type, index and geometry. Use a New-OfficePowerPoint or Open-OfficePowerPoint content block to save completed edits atomically.

## EXAMPLES

### EXAMPLE 1
```powershell
PowerPointPlaceholder -Type Title 'Quarterly review'
```


### EXAMPLE 2
```powershell
Open-OfficePowerPoint -Path ./Review.pptx {
    PowerPointPlaceholder -Slide 2 -Type Body {
        PowerPointParagraph { PowerPointText 'Status: ' -Bold; PowerPointText 'Healthy' }
        PowerPointParagraph 'All regions are operational.'
    }
}
```


## PARAMETERS

### -Content
Native rich paragraphs and runs to place in the existing text box.

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

### -IgnoreMissing
Leave the slide unchanged when the placeholder is absent; the content callback is not run.

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

### -Index
Optional native placeholder index within the selected type.

```yaml
Type: UInt32
Parameter Sets: Text, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Return the same live native text box after editing.

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

### -Slide
Native slide or one-based slide number; omission uses the current slide block.

```yaml
Type: PowerPointSlideSelector
Parameter Sets: Text, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Text
Replacement plain text; an empty string clears the placeholder.

```yaml
Type: String
Parameter Sets: Text
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Type
Placeholder type to update.

```yaml
Type: PowerPointPlaceholderType
Parameter Sets: Text, Content
Aliases: None
Possible values: Title, Body, CenteredTitle, SubTitle, DateAndTime, SlideNumber, Footer, Header, Object, Chart, Table, ClipArt, Diagram, Media, SlideImage, Picture

Required: True
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
