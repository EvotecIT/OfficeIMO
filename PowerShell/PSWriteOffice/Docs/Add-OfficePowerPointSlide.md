---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficePowerPointSlide
## SYNOPSIS
Adds a new slide to a PowerPoint presentation.

## SYNTAX
### Index (Default)
```powershell
Add-OfficePowerPointSlide [[-Content] <scriptblock>] [-Presentation <PowerPointPresentation>] [-Master <int>] [-Layout <int>] [-PassThru] [<CommonParameters>]
```

### Name
```powershell
Add-OfficePowerPointSlide [[-Content] <scriptblock>] -LayoutName <string> [-Presentation <PowerPointPresentation>] [-Master <int>] [-CaseSensitive] [-PassThru] [<CommonParameters>]
```

### Type
```powershell
Add-OfficePowerPointSlide [[-Content] <scriptblock>] -LayoutType <PowerPointSlideLayoutType> [-Presentation <PowerPointPresentation>] [-Master <int>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Creates a slide using OfficeIMO master/layout indexes and can execute nested DSL content.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $ppt = New-OfficePowerPoint -Path .\deck.pptx -NoSave
Add-OfficePowerPointSlide -Presentation $ppt
$ppt | Close-OfficePowerPoint -Save
```

Creates a deck and appends a new slide at the end.

### EXAMPLE 2
```powershell
PS> New-OfficePowerPoint -Path .\deck.pptx { PptSlide { PptTitle -Title 'Status Update' } }
```

Creates a slide and sets the title using DSL aliases.

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

### -Content
Nested DSL content executed within the slide scope.

```yaml
Type: ScriptBlock
Parameter Sets: Index, Name, Type
Aliases: None
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Layout
Layout index to use (matches the template’s built-in layouts).

```yaml
Type: Int32
Parameter Sets: Index
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LayoutName
Layout name to use (case-insensitive by default).

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

### -Presentation
Presentation to update (optional inside New-OfficePowerPoint).

```yaml
Type: PowerPointPresentation
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

- `OfficeIMO.PowerPoint.PowerPointPresentation`

## OUTPUTS

- `None`

## RELATED LINKS

- None
