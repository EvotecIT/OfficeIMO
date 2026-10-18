---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficePowerPointThemeName
## SYNOPSIS
Sets the PowerPoint theme name.

## SYNTAX
### __AllParameterSets
```powershell
Set-OfficePowerPointThemeName -Name <string> [-Presentation <PowerPointPresentation>] [-AllMasters] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Sets the PowerPoint theme name.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficePowerPoint -Path .\Examples\Documents\PowerPointThemeName.pptx {
    Set-OfficePowerPointThemeName -Name 'Service Brief' -AllMasters
    Add-OfficePowerPointSlide -Layout 1 -PassThru | Set-OfficePowerPointSlideTitle -Title 'Named theme'
}
```

Applies a friendly theme name across every master before saving.

## PARAMETERS

### -AllMasters
Apply the name across all slide masters.

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
Theme name to apply.

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
Emit the presentation after update.

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
Presentation to update (optional inside New-OfficePowerPoint).

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

- `OfficeIMO.PowerPoint.PowerPointPresentation`

## RELATED LINKS

- None
