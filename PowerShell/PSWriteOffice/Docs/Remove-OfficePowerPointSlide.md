---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Remove-OfficePowerPointSlide
## SYNOPSIS
Removes a slide by index.

## SYNTAX
### __AllParameterSets
```powershell
Remove-OfficePowerPointSlide -Presentation <PowerPointPresentation> -Index <int> [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Supports -WhatIf/-Confirm thanks to SupportsShouldProcess.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $ppt = New-OfficePowerPoint -Path .\Examples\Documents\PowerPointRemoveSlide.pptx -NoSave
Add-OfficePowerPointSlide -Presentation $ppt -Layout 1
Add-OfficePowerPointSlide -Presentation $ppt -Layout 1
Remove-OfficePowerPointSlide -Presentation $ppt -Index 0 -Confirm:$false
Close-OfficePowerPoint -Presentation $ppt -Save
```

Removes the first slide and saves the updated deck.

## PARAMETERS

### -Index
Zero-based slide index.

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
Presentation to modify.

```yaml
Type: PowerPointPresentation
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

- `None`

## OUTPUTS

- `None`

## RELATED LINKS

- None
