---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficePowerPointThemeColor
## SYNOPSIS
Sets one or more PowerPoint theme colors.

## SYNTAX
### Single (Default)
```powershell
Set-OfficePowerPointThemeColor -Color <PowerPointThemeColor> -Value <string> [-Presentation <PowerPointPresentation>] [-Master <int>] [-AllMasters] [-PassThru] [<CommonParameters>]
```

### Multiple
```powershell
Set-OfficePowerPointThemeColor -Colors <hashtable> [-Presentation <PowerPointPresentation>] [-Master <int>] [-AllMasters] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Sets one or more PowerPoint theme colors.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Set-OfficePowerPointThemeColor -Presentation $ppt -Color Accent1 -Value '#C00000'
```

Updates Accent1 on the default master.

### EXAMPLE 2
```powershell
PS> Set-OfficePowerPointThemeColor -Presentation $ppt -Colors @{ Accent1 = '#C00000'; Accent2 = '#00B0F0' } -AllMasters
```

Applies multiple theme colors to every master in the presentation.

## PARAMETERS

### -AllMasters
Apply the changes across all slide masters.

```yaml
Type: SwitchParameter
Parameter Sets: Single, Multiple
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Color
Theme color to update.

```yaml
Type: PowerPointThemeColor
Parameter Sets: Single
Aliases: None
Possible values: Dark1, Light1, Dark2, Light2, Accent1, Accent2, Accent3, Accent4, Accent5, Accent6, Hyperlink, FollowedHyperlink

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Colors
Hashtable of theme color names to named or hexadecimal color values.

```yaml
Type: Hashtable
Parameter Sets: Multiple
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Master
Slide master index to update when not using AllMasters.

```yaml
Type: Int32
Parameter Sets: Single, Multiple
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the presentation after update.

```yaml
Type: SwitchParameter
Parameter Sets: Single, Multiple
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
Parameter Sets: Single, Multiple
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Value
Named or hexadecimal color value.

```yaml
Type: String
Parameter Sets: Single
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

- `OfficeIMO.PowerPoint.PowerPointPresentation`

## RELATED LINKS

- None
