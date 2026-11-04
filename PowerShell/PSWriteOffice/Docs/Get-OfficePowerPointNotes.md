---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficePowerPointNotes
## SYNOPSIS
Reads speaker notes without creating empty notes parts.

## SYNTAX
### Document (Default)
```powershell
Get-OfficePowerPointNotes [-Document <PowerPointDocument>] [-Slide <PowerPointSlideSelector>] [-IncludeEmpty] [<CommonParameters>]
```

## DESCRIPTION
Reads speaker notes without creating empty notes parts.

## EXAMPLES

### EXAMPLE 1
```powershell
Get-OfficePowerPointNotes -Document $deck -IncludeEmpty | Select-Object SlideIndex, Text
```


### EXAMPLE 2
```powershell
Get-OfficePowerPointNotes -Document $deck -Slide 3
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
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -IncludeEmpty
Gets or sets whether slides without notes are included.

```yaml
Type: SwitchParameter
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Slide
Gets or sets a native slide object or one-based slide position.

```yaml
Type: PowerPointSlideSelector
Parameter Sets: Document
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

## OUTPUTS

- `PSWriteOffice.Services.PowerPoint.PowerPointNotesInfo`: PowerShell-friendly description of slide speaker notes.

## RELATED LINKS

- None
