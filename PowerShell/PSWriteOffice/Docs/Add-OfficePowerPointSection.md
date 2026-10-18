---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficePowerPointSection
## SYNOPSIS
Adds a section to a PowerPoint presentation.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficePowerPointSection -Name <string> [-Presentation <PowerPointPresentation>] [-StartSlideIndex <Int32>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Creates a new section starting at the requested slide index or at the current slide inside the DSL.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficePowerPoint -Path .\Examples\Documents\PowerPointSections.pptx {
    Add-OfficePowerPointSlide -Layout 1 -PassThru | Set-OfficePowerPointSlideTitle -Title 'Overview'
    Add-OfficePowerPointSlide -Layout 1 -PassThru | Set-OfficePowerPointSlideTitle -Title 'Results'
    Add-OfficePowerPointSection -Name 'Results' -StartSlideIndex 1
}
```

Creates a section named Results starting at the second slide.

## PARAMETERS

### -Name
Name of the section to add.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
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

### -StartSlideIndex
Zero-based slide index where the section should start.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.PowerPoint.PowerPointPresentation`

## OUTPUTS

- `OfficeIMO.PowerPoint.PowerPointSectionInfo`

## RELATED LINKS

- None
