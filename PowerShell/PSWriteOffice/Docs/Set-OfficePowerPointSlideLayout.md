---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficePowerPointSlideLayout
## SYNOPSIS
Changes the layout used by a slide.

## SYNTAX
### Index (Default)
```powershell
Set-OfficePowerPointSlideLayout -Layout <int> [-Slide <PowerPointSlide>] [-Master <int>] [-PassThru] [<CommonParameters>]
```

### Name
```powershell
Set-OfficePowerPointSlideLayout -LayoutName <string> [-Slide <PowerPointSlide>] [-Master <int>] [-CaseSensitive] [-PassThru] [<CommonParameters>]
```

### Type
```powershell
Set-OfficePowerPointSlideLayout -LayoutType <PowerPointSlideLayoutType> [-Slide <PowerPointSlide>] [-Master <int>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Changes the layout used by a slide.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Get-OfficePowerPointSlide -Presentation $ppt -Index 0 | Set-OfficePowerPointSlideLayout -LayoutName 'Title and Content'
```

Updates the slide to use the requested layout.

## PARAMETERS

### -CaseSensitive
Use case-sensitive matching for layout names.

```yaml
Type: SwitchParameter
Parameter Sets: Name
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Layout
Layout index to use.

```yaml
Type: Int32
Parameter Sets: Index
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LayoutName
Layout name to use.

```yaml
Type: String
Parameter Sets: Name
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LayoutType
Layout type to use.

```yaml
Type: PowerPointSlideLayoutType
Parameter Sets: Type
Aliases: None
Possible values: Title, Text, TwoColumnText, Table, TextAndChart, ChartAndText, Diagram, Chart, TextAndClipArt, ClipArtAndText, TitleOnly, Blank, TextAndObject, ObjectAndText, ObjectOnly, Object, TextAndMedia, MediaAndText, ObjectOverText, TextOverObject, TextAndTwoObjects, TwoObjectsAndText, TwoObjectsOverText, FourObjects, VerticalText, ClipArtAndVerticalText, VerticalTitleAndText, VerticalTitleAndTextOverChart, TwoObjects, ObjectAndTwoObjects, TwoObjectsAndObject, Custom, SectionHeader, TwoTextAndTwoObjects, ObjectText, PictureText

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Master
Slide master index to use.

```yaml
Type: Int32
Parameter Sets: Index, Name, Type
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the object created or changed by the command.

```yaml
Type: SwitchParameter
Parameter Sets: Index, Name, Type
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Slide
Slide to update (optional inside a slide DSL scope).

```yaml
Type: PowerPointSlide
Parameter Sets: Index, Name, Type
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.PowerPoint.PowerPointSlide`

## OUTPUTS

- `OfficeIMO.PowerPoint.PowerPointSlide`

## RELATED LINKS

- None
