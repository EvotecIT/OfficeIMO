---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Copy-OfficePowerPointSlide
## SYNOPSIS
Copies a native slide within its document, preserving its content and relationships.

## SYNTAX
### __AllParameterSets
```powershell
Copy-OfficePowerPointSlide [-Slide] <PowerPointSlideSelector> [-Document <PowerPointDocument>] [-Position <Int32>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Copies a native slide within its document, preserving its content and relationships.

## EXAMPLES

### EXAMPLE 1
```powershell
Copy-OfficePowerPointSlide -Document $deck -Slide 1 -Position 3 -PassThru
```


### EXAMPLE 2
```powershell
$slide | Copy-OfficePowerPointSlide -PassThru
```


## PARAMETERS

### -Document
Gets or sets the document for numeric selection; omission uses the enclosing scope or native slide owner.

```yaml
Type: PowerPointDocument
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
Gets or sets whether to return the new native slide.

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

### -Position
Gets or sets the copy's one-based destination position; omission inserts it immediately after the source.

```yaml
Type: Int32
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
Gets or sets a native slide or its one-based position.

```yaml
Type: PowerPointSlideSelector
Parameter Sets: __AllParameterSets
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

- `PSWriteOffice.Foundation.PowerPointSlideSelector`: A live slide or a one-based slide position in a document.

## OUTPUTS

- `OfficeIMO.PowerPoint.PowerPointSlide`

## RELATED LINKS

- None
