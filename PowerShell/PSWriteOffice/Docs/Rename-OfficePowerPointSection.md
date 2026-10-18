---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Rename-OfficePowerPointSection
## SYNOPSIS
Renames a PowerPoint section.

## SYNTAX
### __AllParameterSets
```powershell
Rename-OfficePowerPointSection -Name <string> -NewName <string> [-Presentation <PowerPointPresentation>] [-CaseSensitive] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Renames a PowerPoint section.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $ppt = New-OfficePowerPoint -Path .\Examples\Documents\PowerPointRenameSection.pptx -NoSave
Add-OfficePowerPointSlide -Presentation $ppt -Layout 1
Add-OfficePowerPointSection -Presentation $ppt -Name 'Results' -StartSlideIndex 0
Rename-OfficePowerPointSection -Presentation $ppt -Name 'Results' -NewName 'Deep Dive' -PassThru
$ppt | Close-OfficePowerPoint -Save
```

Renames the first matching section and returns the updated section metadata.

## PARAMETERS

### -CaseSensitive
Use case-sensitive matching for the existing section name.

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
Existing section name.

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

### -NewName
New section name.

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
Emit the renamed section instead of no output.

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

- `OfficeIMO.PowerPoint.PowerPointSectionInfo`
- `System.Boolean`

## RELATED LINKS

- None
