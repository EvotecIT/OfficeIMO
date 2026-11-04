---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficePowerPointLayout
## SYNOPSIS
Lists template layouts from all masters, or one selected master.

## SYNTAX
### Document (Default)
```powershell
Get-OfficePowerPointLayout [[-Document] <PowerPointDocument>] [-Master <Int32>] [<CommonParameters>]
```

## DESCRIPTION
Lists template layouts from all masters, or one selected master.

## EXAMPLES

### EXAMPLE 1
```powershell
Get-OfficePowerPointLayout -Document $deck
```


### EXAMPLE 2
```powershell
New-OfficePowerPoint -Path ./Review.pptx {
    $layout = @(Get-OfficePowerPointLayout -Master 1)[0]
    PowerPointSlide -Title 'Review' -TemplateLayout $layout
}
```


## PARAMETERS

### -Document
Gets or sets the live document; omission uses the enclosing content scope.

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

### -Master
Gets or sets a one-based master position; omission lists all masters.

```yaml
Type: Int32
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

- `OfficeIMO.PowerPoint.PowerPointSlideLayoutInfo`

## RELATED LINKS

- None
