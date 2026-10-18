---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficePowerPointInspection
## SYNOPSIS
Runs package, preflight, accessibility, feature, review, animation, signature, and visual inspections.

## SYNTAX
### Path (Default)
```powershell
Get-OfficePowerPointInspection [-Path] <string> [-Options <PowerPointInspectionOptions>] [<CommonParameters>]
```

### Presentation
```powershell
Get-OfficePowerPointInspection -Presentation <PowerPointPresentation> [-Options <PowerPointInspectionOptions>] [<CommonParameters>]
```

## DESCRIPTION
Runs package, preflight, accessibility, feature, review, animation, signature, and visual inspections.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Get-OfficePowerPointInspection -Path .\Deck.pptx
```

Returns one coherent inspection report over the same presentation model used for editing.

## PARAMETERS

### -Options
Optional report selection and inspection policies.

```yaml
Type: PowerPointInspectionOptions
Parameter Sets: Path, Presentation
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Path to the presentation.

```yaml
Type: String
Parameter Sets: Path
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Presentation
Open presentation instance.

```yaml
Type: PowerPointPresentation
Parameter Sets: Presentation
Aliases: None
Possible values:

Required: True
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

- `OfficeIMO.PowerPoint.PowerPointInspectionReport`

## RELATED LINKS

- None
