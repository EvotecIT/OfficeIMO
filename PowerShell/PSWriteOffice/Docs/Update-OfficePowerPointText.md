---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Update-OfficePowerPointText
## SYNOPSIS
Replaces text in the selected slide or document, including tables by default.

## SYNTAX
### Document (Default)
```powershell
Update-OfficePowerPointText [-OldValue] <string> [-NewValue] <string> [-Document <PowerPointDocument>] [-Slide <PowerPointSlideSelector>] [-NoTables] [-IncludeNotes] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Uses the current slide inside PowerPointSlide and the current document inside an open document block. Notes are included only with IncludeNotes.

## EXAMPLES

### EXAMPLE 1
```powershell
Open-OfficePowerPoint -Path .\Account-Review.pptx {
    Update-OfficePowerPointText '{{Customer}}' 'Northwind Traders' -IncludeNotes
}
```


### EXAMPLE 2
```powershell
$count = Update-OfficePowerPointText -Document $deck -Slide 2 -OldValue 'FY24' -NewValue 'FY25' -NoTables -PassThru
```


## PARAMETERS

### -Document
Document to update or to resolve a one-based slide position against; omission uses the current block.

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

### -IncludeNotes
Include speaker notes in the replacement operation.

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

### -NewValue
Replacement text; null clears each match.

```yaml
Type: String
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoTables
Leave table cell text unchanged.

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

### -OldValue
Text to find.

```yaml
Type: String
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the object created or changed by the command.

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
Native slide or one-based slide position; omission updates the current slide or whole document.

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

- `System.Int32`

## RELATED LINKS

- None
