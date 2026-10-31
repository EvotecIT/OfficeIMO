---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficePowerPointSlide
## SYNOPSIS
Gets editable slides from a live document, or detached slide summaries from files.

## SYNTAX
### Document (Default)
```powershell
Get-OfficePowerPointSlide [[-Document] <PowerPointDocument>] [-Slide <int[]>] [-Summary] [<CommonParameters>]
```

### Path
```powershell
Get-OfficePowerPointSlide [-Path] <string[]> [-Slide <int[]>] [-Summary] [<CommonParameters>]
```

### LiteralPath
```powershell
Get-OfficePowerPointSlide -LiteralPath <string[]> [-Slide <int[]>] [-Summary] [<CommonParameters>]
```

## DESCRIPTION
Gets editable slides from a live document, or detached slide summaries from files.

## EXAMPLES

### EXAMPLE 1
```powershell
Get-ChildItem .\Decks -Filter *.pptx | Get-OfficePowerPointSlide -Summary | Select-Object Path, SlideIndex, Title
```


### EXAMPLE 2
```powershell
Get-OfficePowerPointSlide -Document $deck -Slide 3
```


## PARAMETERS

### -Document
Gets or sets a live document; omission uses the enclosing document block.

```yaml
Type: PowerPointDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -LiteralPath
Gets or sets literal file paths, including FileInfo pipeline input.

```yaml
Type: String[]
Parameter Sets: LiteralPath
Aliases: PSPath, FullName
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -Path
Gets or sets file paths with provider wildcard expansion.

```yaml
Type: String[]
Parameter Sets: Path
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -Slide
Gets or sets one-based slide positions; omission selects every slide.

```yaml
Type: Int32[]
Parameter Sets: Document, Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Summary
Gets or sets whether to return summaries; required when inspecting files.

```yaml
Type: SwitchParameter
Parameter Sets: Document, Path, LiteralPath
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

- `OfficeIMO.PowerPoint.PowerPointDocument`
- `System.String[]`

## OUTPUTS

- `OfficeIMO.PowerPoint.PowerPointSlide`
- `PSWriteOffice.Services.PowerPoint.PowerPointSlideSummaryInfo`: PowerShell-friendly summary of a PowerPoint slide.

## RELATED LINKS

- None
