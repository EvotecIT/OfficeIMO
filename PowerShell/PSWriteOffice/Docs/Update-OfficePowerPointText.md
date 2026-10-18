---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Update-OfficePowerPointText
## SYNOPSIS
Replaces text in a PowerPoint slide or presentation.

## SYNTAX
### Auto (Default)
```powershell
Update-OfficePowerPointText -OldValue <string> -NewValue <string> [-IncludeTables <bool>] [-IncludeNotes] [-PassThru] [<CommonParameters>]
```

### Presentation
```powershell
Update-OfficePowerPointText -OldValue <string> -NewValue <string> [-Presentation <PowerPointPresentation>] [-IncludeTables <bool>] [-IncludeNotes] [-PassThru] [<CommonParameters>]
```

### Slide
```powershell
Update-OfficePowerPointText -OldValue <string> -NewValue <string> [-Slide <PowerPointSlide>] [-IncludeTables <bool>] [-IncludeNotes] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Can replace text in text boxes, tables, and optionally notes using the OfficeIMO text replacement helpers.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $ppt = New-OfficePowerPoint -Path .\Examples\Documents\PowerPointUpdateText.pptx -NoSave
$slide = Add-OfficePowerPointSlide -Presentation $ppt -Layout 1 -PassThru
Add-OfficePowerPointTextBox -Slide $slide -Text 'FY24 summary'
Set-OfficePowerPointNotes -Slide $slide -Text 'Mention FY24 assumptions.'
$count = Update-OfficePowerPointText -Presentation $ppt -OldValue 'FY24' -NewValue 'FY25' -IncludeNotes -PassThru
$ppt | Close-OfficePowerPoint -Save
```

Replaces matching text throughout the presentation and notes, returning the replacement count.

## PARAMETERS

### -IncludeNotes
Include notes text in the replacement operation.

```yaml
Type: SwitchParameter
Parameter Sets: Auto, Presentation, Slide
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IncludeTables
Include table cells in the replacement operation.

```yaml
Type: Boolean
Parameter Sets: Auto, Presentation, Slide
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NewValue
Replacement text.

```yaml
Type: String
Parameter Sets: Auto, Presentation, Slide
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OldValue
Text to find.

```yaml
Type: String
Parameter Sets: Auto, Presentation, Slide
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the object created or changed by the command.

```yaml
Type: SwitchParameter
Parameter Sets: Auto, Presentation, Slide
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Presentation
Presentation to update.

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
Slide to update.

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

- `OfficeIMO.PowerPoint.PowerPointPresentation`
- `OfficeIMO.PowerPoint.PowerPointSlide`

## OUTPUTS

- `System.Int32`

## RELATED LINKS

- None
