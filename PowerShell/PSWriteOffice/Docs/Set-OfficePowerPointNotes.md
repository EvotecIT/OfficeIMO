---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Set-OfficePowerPointNotes
## SYNOPSIS
Sets speaker notes on a native slide or in the current slide content block.

## SYNTAX
### __AllParameterSets
```powershell
Set-OfficePowerPointNotes [-Text] <string> [-Slide <PowerPointSlideSelector>] [-Document <PowerPointPresentation>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Sets speaker notes on a native slide or in the current slide content block.

## EXAMPLES

### EXAMPLE 1
```powershell
New-OfficePowerPoint -Path ./Review.pptx {
    PowerPointSlide -Title 'Service review' {
        PowerPointParagraph 'All regions healthy'
        PowerPointNotes 'Pause for questions before discussing capacity.'
    }
}
```


### EXAMPLE 2
```powershell
Get-OfficePowerPointSlide -Document $deck -Slide 2 |
    Set-OfficePowerPointNotes -Text 'Explain the margin before revenue.'
Save-OfficePowerPoint -Document $deck
```


## PARAMETERS

### -Document
Gets or sets the document for numeric selection; omission uses the enclosing document block.

```yaml
Type: PowerPointPresentation
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
Gets or sets whether to return the selected native slide after editing.

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

### -Slide
Gets or sets a native slide or one-based position; omission uses the current slide content block.

```yaml
Type: PowerPointSlideSelector
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Text
Gets or sets the notes text; an empty string clears the notes.

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

- `PSWriteOffice.Foundation.PowerPointSlideSelector`: A live slide or a one-based slide position in a document.

## OUTPUTS

- `OfficeIMO.PowerPoint.PowerPointSlide`

## RELATED LINKS

- None
