---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficePowerPointSlideSummary
## SYNOPSIS
Reads high-level slide summaries from a presentation.

## SYNTAX
### Slide (Default)
```powershell
Get-OfficePowerPointSlideSummary [-Slide <PowerPointSlide>] [<CommonParameters>]
```

### Presentation
```powershell
Get-OfficePowerPointSlideSummary [-Presentation <PowerPointPresentation>] [-Index <Int32>] [<CommonParameters>]
```

## DESCRIPTION
Returns title, notes metadata, layout metadata, and content counts for each slide.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Get-OfficePowerPointSlideSummary -Presentation $ppt
```

Returns one summary object per slide.

### EXAMPLE 2
```powershell
PS> Get-OfficePowerPointSlide -Presentation $ppt -Index 0 | Get-OfficePowerPointSlideSummary
```

Returns the summary for the selected slide.

## PARAMETERS

### -Index
Optional zero-based slide index when reading from a presentation.

```yaml
Type: Int32
Parameter Sets: Presentation
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Presentation
Presentation whose slides should be summarized.

```yaml
Type: PowerPointPresentation
Parameter Sets: Presentation
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Slide
Slide to inspect (optional inside the DSL).

```yaml
Type: PowerPointSlide
Parameter Sets: Slide
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

- `OfficeIMO.PowerPoint.PowerPointSlide`
- `OfficeIMO.PowerPoint.PowerPointPresentation`

## OUTPUTS

- `PSWriteOffice.Services.PowerPoint.PowerPointSlideSummaryInfo`: PowerShell-friendly summary of a PowerPoint slide.

## RELATED LINKS

- None
