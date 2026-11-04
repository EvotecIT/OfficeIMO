---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficePowerPointSection
## SYNOPSIS
Starts a named section at an existing, current or next appended slide.

## SYNTAX
### Name (Default)
```powershell
Add-OfficePowerPointSection [-Name] <string> [-Document <PowerPointDocument>] [-Slide <PowerPointSlideSelector>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Starts a named section at an existing, current or next appended slide.

## EXAMPLES

### EXAMPLE 1
```powershell
New-OfficePowerPoint -Path ./Results.pptx {
    PowerPointSection 'Results'
    PowerPointSlide -Title 'Revenue' { PowerPointParagraph 'Revenue increased.' }
    PowerPointSlide -Title 'Margin' { PowerPointParagraph 'Margin improved.' }
}
```


### EXAMPLE 2
```powershell
Add-OfficePowerPointSection -Document $deck -Name Appendix -Slide 9
```


## PARAMETERS

### -Document
Gets or sets the document; omission uses the enclosing scope or selected native slide's owner.

```yaml
Type: PowerPointDocument
Parameter Sets: Name
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Name
Gets or sets the section name.

```yaml
Type: String
Parameter Sets: Name
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Gets or sets whether to return section metadata, including its stable native identifier.

```yaml
Type: SwitchParameter
Parameter Sets: Name
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Slide
Gets or sets a native slide or one-based slide position; omission uses the current or next slide.

```yaml
Type: PowerPointSlideSelector
Parameter Sets: Name
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

- `OfficeIMO.PowerPoint.PowerPointSectionInfo`

## RELATED LINKS

- None
