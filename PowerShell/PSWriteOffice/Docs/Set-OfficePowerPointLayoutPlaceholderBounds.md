---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficePowerPointLayoutPlaceholderBounds
## SYNOPSIS
Sets layout placeholder bounds for a slide layout.

## SYNTAX
### __AllParameterSets
```powershell
Set-OfficePowerPointLayoutPlaceholderBounds -Layout <int> -PlaceholderType <PowerPointPlaceholderType> -Left <double> -Top <double> -Width <double> -Height <double> [-Presentation <PowerPointPresentation>] [-Master <int>] [-Index <UInt32>] [-CreateIfMissing] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Sets layout placeholder bounds for a slide layout.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Set-OfficePowerPointLayoutPlaceholderBounds -Presentation $ppt -Master 0 -Layout 1 -PlaceholderType Title -Left 40 -Top 20 -Width 500 -Height 120
```

Moves/resizes the Title placeholder on the layout.

### EXAMPLE 2
```powershell
PS> New-OfficePowerPoint -Path .\deck.pptx {
  $layout = Get-OfficePowerPointLayout | Select-Object -First 1
  Set-OfficePowerPointLayoutPlaceholderBounds -Master $layout.MasterIndex -Layout $layout.LayoutIndex -PlaceholderType Title -Left 40 -Top 20 -Width 500 -Height 120
}
```

Uses the DSL context to resolve the presentation.

## PARAMETERS

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

### -Height
Height in points.

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
Left position in points.

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

### -Top
Top position in points.

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

### -Width
Width in points.

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
