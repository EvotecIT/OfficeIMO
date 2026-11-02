---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Remove-OfficePowerPointSlide
## SYNOPSIS
Removes selected native slides or one-based slide positions.

## SYNTAX
### Slide (Default)
```powershell
Remove-OfficePowerPointSlide [-Slide] <PowerPointSlideSelector[]> [-PassThru] [-Document <PowerPointDocument>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Removes selected native slides or one-based slide positions.

## EXAMPLES

### EXAMPLE 1
```powershell
Remove-OfficePowerPointSlide -Document $deck -Slide 1, 3 -Confirm:$false
```


### EXAMPLE 2
```powershell
$slide | Remove-OfficePowerPointSlide -WhatIf
```


## PARAMETERS

### -Document
Gets or sets the document for numeric selection; omission uses the enclosing scope or native slide owner.

```yaml
Type: PowerPointDocument
Parameter Sets: Slide
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Return the live document after at least one slide is removed.

```yaml
Type: SwitchParameter
Parameter Sets: Slide
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Slide
Gets or sets native slides or their one-based positions; each argument array is resolved before removal.

```yaml
Type: PowerPointSlideSelector[]
Parameter Sets: Slide
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `PSWriteOffice.Foundation.PowerPointSlideSelector[]`: A live slide or a one-based slide position in a document.

## OUTPUTS

- `OfficeIMO.PowerPoint.PowerPointDocument`

## RELATED LINKS

- None
