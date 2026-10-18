---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficePowerPointLayoutPlaceholderTextMargins
## SYNOPSIS
Sets layout placeholder text margins for a slide layout (points).

## SYNTAX
### __AllParameterSets
```powershell
Set-OfficePowerPointLayoutPlaceholderTextMargins -Layout <int> -PlaceholderType <PowerPointPlaceholderType> -Left <double> -Top <double> -Right <double> -Bottom <double> [-Presentation <PowerPointPresentation>] [-Master <int>] [-Index <UInt32>] [-CreateIfMissing] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Sets layout placeholder text margins for a slide layout (points).

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Set-OfficePowerPointLayoutPlaceholderTextMargins -Presentation $ppt -Master 0 -Layout 1 -PlaceholderType Title -Left 12 -Top 8 -Right 12 -Bottom 8
```

Updates the text margins on the layout placeholder.

### EXAMPLE 2
```powershell
PS> New-OfficePowerPoint -Path .\deck.pptx {
  $layout = Get-OfficePowerPointLayout | Select-Object -First 1
  Set-OfficePowerPointLayoutPlaceholderTextMargins -Master $layout.MasterIndex -Layout $layout.LayoutIndex -PlaceholderType Title -Left 12 -Top 8 -Right 12 -Bottom 8
}
```

Uses the DSL context to resolve the presentation.

## PARAMETERS

### -Bottom
Bottom margin in points.

```yaml
Type: Double
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -CreateIfMissing
Create the placeholder if it is missing.

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

### -Index
Optional placeholder index.

```yaml
Type: UInt32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Layout
Layout index within the master.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Left
Left margin in points.

```yaml
Type: Double
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Master
Slide master index.

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

### -PassThru
Emit the placeholder textbox after update.

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

### -PlaceholderType
Placeholder type to target.

```yaml
Type: PowerPointPlaceholderType
Parameter Sets: __AllParameterSets
Aliases: Type
Possible values: Title, Body, CenteredTitle, SubTitle, DateAndTime, SlideNumber, Footer, Header, Object, Chart, Table, ClipArt, Diagram, Media, SlideImage, Picture

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Presentation
Presentation to update (optional inside DSL).

```yaml
Type: PowerPointPresentation
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Right
Right margin in points.

```yaml
Type: Double
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Top
Top margin in points.

```yaml
Type: Double
Parameter Sets: __AllParameterSets
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

- `OfficeIMO.PowerPoint.PowerPointPresentation`

## OUTPUTS

- `OfficeIMO.PowerPoint.PowerPointTextBox`

## RELATED LINKS

- None
