---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficePowerPointPlanLogoWall
## SYNOPSIS
Adds a semantic logo/proof-wall slide to a PowerPoint deck plan.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficePowerPointPlanLogoWall [-Title] <string> -Logos <Object[]> [-Plan <PowerPointDeckPlan>] [-Subtitle <string>] [-Seed <string>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a semantic logo/proof-wall slide to a PowerPoint deck plan.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $logos = @(
    @{ Name = 'Directory'; Subtitle = 'Identity platform'; AccentColor = '#2563EB' }
    @{ Name = 'Mail'; Subtitle = 'Messaging platform'; AccentColor = '#0F766E' }
)
New-OfficePowerPointDeckPlan {
    Add-OfficePowerPointPlanLogoWall -Title 'Systems covered' -Subtitle 'Representative services' -Logos $logos
}
```

Adds a semantic logo/proof-wall slide to the plan.

## PARAMETERS

### -Logos
Objects with Name, optional Subtitle, ImagePath, and AccentColor properties.

```yaml
Type: Object[]
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
Emit the updated plan.

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

### -Plan
Plan to update. Optional inside New-OfficePowerPointDeckPlan.

```yaml
Type: PowerPointDeckPlan
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Seed
Stable seed for deterministic visual selection.

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

### -Subtitle
Optional slide subtitle.

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

### -Title
Slide title.

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

- `OfficeIMO.PowerPoint.PowerPointDeckPlan`

## OUTPUTS

- `OfficeIMO.PowerPoint.PowerPointDeckPlan`

## RELATED LINKS

- None
