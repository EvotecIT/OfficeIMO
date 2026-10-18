---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Copy-OfficePowerPointSlide
## SYNOPSIS
Copies an existing slide within a PowerPoint presentation.

## SYNTAX
### __AllParameterSets
```powershell
Copy-OfficePowerPointSlide -Index <int> [-Presentation <PowerPointPresentation>] [-InsertAt <Int32>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Uses OfficeIMO slide duplication so charts, notes, and shapes are preserved.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficePowerPoint -Path .\Examples\Documents\PowerPointCopySlide.pptx {
    $slide = Add-OfficePowerPointSlide -Layout 1 -PassThru
    Set-OfficePowerPointSlideTitle -Slide $slide -Title 'Original'
    $copy = Copy-OfficePowerPointSlide -Index 0 -PassThru
    Set-OfficePowerPointSlideTitle -Slide $copy -Title 'Copied appendix'
}
```

Duplicates a slide and updates the copied slide title.

## PARAMETERS

### -Index
Zero-based slide index to duplicate.

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

### -InsertAt
Optional target index for the duplicate; omit to insert after the source slide.

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
Emit the object created or changed by the command.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.PowerPoint.PowerPointPresentation`

## OUTPUTS

- `OfficeIMO.PowerPoint.PowerPointSlide`

## RELATED LINKS

- None
