---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficePowerPointNotes
## SYNOPSIS
Sets speaker notes for a PowerPoint slide.

## SYNTAX
### __AllParameterSets
```powershell
Set-OfficePowerPointNotes [-Text] <string> [-Slide <PowerPointSlide>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Sets speaker notes for a PowerPoint slide.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficePowerPoint -Path .\Examples\Documents\PowerPointNotes.pptx {
    $slide = Add-OfficePowerPointSlide -Layout 1 -PassThru
    Set-OfficePowerPointSlideTitle -Slide $slide -Title 'Executive summary'
    Set-OfficePowerPointNotes -Slide $slide -Text 'Keep this slide under five minutes and focus on decisions.'
}
```

Writes speaker notes to a generated slide.

## PARAMETERS

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

### -Slide
Slide whose notes should be updated (optional inside DSL).

```yaml
Type: PowerPointSlide
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Text
Notes text to apply.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
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
