---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficePowerPointPlanCaseStudy
## SYNOPSIS
Adds a semantic case-study slide to a PowerPoint deck plan.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficePowerPointPlanCaseStudy [-Title] <string> -Sections <Object[]> [-Plan <PowerPointDeckPlan>] [-Metrics <Object[]>] [-Seed <string>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a semantic case-study slide to a PowerPoint deck plan.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $sections = @(
    @{ Heading = 'Challenge'; Body = 'Manual reports took too long to produce.' }
    @{ Heading = 'Outcome'; Body = 'Automated generation made the review repeatable.' }
)
$metrics = @(
    @{ Value = '4h'; Label = 'saved each cycle' }
    @{ Value = '0'; Label = 'manual copy steps' }
)
New-OfficePowerPointDeckPlan {
    Add-OfficePowerPointPlanCaseStudy -Title 'Automation impact' -Sections $sections -Metrics $metrics
}
```

Adds a proof-oriented case-study slide to the plan.

## PARAMETERS

### -Metrics
Objects with Value and Label/Name/Title properties.

```yaml
Type: Object[]
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
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

### -Sections
Objects with Heading/Title and Body/Description/Text properties.

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
