---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficePowerPointSection
## SYNOPSIS
Gets PowerPoint sections from a presentation.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficePowerPointSection [-Presentation <PowerPointPresentation>] [-Name <string>] [-CaseSensitive] [<CommonParameters>]
```

## DESCRIPTION
Returns OfficeIMO section metadata so scripts can inspect section names and slide ranges.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $ppt = New-OfficePowerPoint -Path .\Examples\Documents\PowerPointSectionsRead.pptx -NoSave
Add-OfficePowerPointSlide -Presentation $ppt -Layout 1
Add-OfficePowerPointSection -Presentation $ppt -Name 'Appendix' -StartSlideIndex 0
Get-OfficePowerPointSection -Presentation $ppt | Select-Object Name, FirstSlideIndex, SlideCount
$ppt | Close-OfficePowerPoint
```

Returns section information including section names and slide indexes.

## PARAMETERS

### -CaseSensitive
Use case-sensitive matching for section names.

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

### -Name
Optional section name filter.

```yaml
Type: String
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
Presentation to inspect (optional inside DSL).

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

- `OfficeIMO.PowerPoint.PowerPointSectionInfo`

## RELATED LINKS

- None
