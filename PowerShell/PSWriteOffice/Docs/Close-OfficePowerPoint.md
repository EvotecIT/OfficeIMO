---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Close-OfficePowerPoint
## SYNOPSIS
Closes a PowerPoint presentation and optionally saves it.

## SYNTAX
### __AllParameterSets
```powershell
Close-OfficePowerPoint -Presentation <PowerPointPresentation> [-Save] [-Path <string>] [-Open] [-Password <string>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Provides a cmdlet wrapper so PowerShell scripts do not need to call Dispose directly.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $ppt = Get-OfficePowerPoint -Path .\deck.pptx; Close-OfficePowerPoint -Presentation $ppt
```

Releases the loaded presentation instance.

### EXAMPLE 2
```powershell
PS> Close-OfficePowerPoint -Presentation $ppt -Save -Open
```

Saves the presentation, opens it in PowerPoint, and releases the object.

## PARAMETERS

### -Open
Open the presentation after saving. Requires -Save or -Path.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: Show
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Password
Password used to save the presentation as an encrypted package.

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

### -Path
Optional target path when saving.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: FilePath
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Presentation
Presentation to close.

```yaml
Type: PowerPointPresentation
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Save
Persist changes before closing.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.PowerPoint.PowerPointPresentation`

## OUTPUTS

- `None`

## RELATED LINKS

- None
