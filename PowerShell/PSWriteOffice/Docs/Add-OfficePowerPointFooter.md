---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficePowerPointFooter
## SYNOPSIS
Sets a deck or logical-slide footer slot using text or native inline content.

## SYNTAX
### Text (Default)
```powershell
Add-OfficePowerPointFooter [-Text] <string> [-Alignment <OfficeAlignment>] [<CommonParameters>]
```

### Content
```powershell
Add-OfficePowerPointFooter [-Content] <scriptblock> [-Alignment <OfficeAlignment>] [<CommonParameters>]
```

## DESCRIPTION
At the New-/Open- block level the slot applies throughout the deck, including existing slides. Inside PowerPointSlide it overrides that slot for the slide and its continuations. The content block runs once; repeated footers retain native styles, fields and links.

## EXAMPLES

### EXAMPLE 1
```powershell
New-OfficePowerPoint -Path ./Review.pptx {
    PowerPointFooter 'Contoso | Confidential'
    PowerPointFooter -Alignment Right { PowerPointSlideNumber -IncludeTotal }
    PowerPointSlide -Title 'Service review'
}
```


## PARAMETERS

### -Alignment
Left, Center or Right slot; defaults to Left.

```yaml
Type: OfficeAlignment
Parameter Sets: Text, Content
Aliases: None
Possible values: Left, Center, Right, Justify, Start, End

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Content
Inline content written with PowerPointText and PowerPointSlideNumber.

```yaml
Type: ScriptBlock
Parameter Sets: Content
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Text
Plain text for the footer slot.

```yaml
Type: String
Parameter Sets: Text
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

- `None`

## OUTPUTS

- `None`

## RELATED LINKS

- None
