---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficePowerPointNotes
## SYNOPSIS
Reads speaker notes from a slide or presentation.

## SYNTAX
### Slide (Default)
```powershell
Get-OfficePowerPointNotes [-Slide <PowerPointSlide>] [-IncludeEmpty] [<CommonParameters>]
```

### Presentation
```powershell
Get-OfficePowerPointNotes [-Presentation <PowerPointPresentation>] [-Index <Int32>] [-IncludeEmpty] [<CommonParameters>]
```

## DESCRIPTION
Returns note metadata without creating empty notes parts on slides that do not already have them.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Get-OfficePowerPointSlide -Presentation $ppt -Index 0 | Get-OfficePowerPointNotes
```

Returns the notes text and metadata for the selected slide.

### EXAMPLE 2
```powershell
PS> Get-OfficePowerPointNotes -Presentation $ppt -IncludeEmpty
```

Lists slide indexes together with note text, including slides that have no notes yet.

## PARAMETERS

### -IncludeEmpty
Include slides that do not currently have speaker notes.

```yaml
Type: SwitchParameter
Parameter Sets: Slide, Presentation
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

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
Presentation whose slides should be inspected.

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

- `PSWriteOffice.Services.PowerPoint.PowerPointNotesInfo`: PowerShell-friendly description of slide speaker notes.

## RELATED LINKS

- None
