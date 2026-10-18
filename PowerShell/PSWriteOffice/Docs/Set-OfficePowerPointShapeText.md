---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficePowerPointShapeText
## SYNOPSIS
Sets text on an existing PowerPoint text box.

## SYNTAX
### Text (Default)
```powershell
Set-OfficePowerPointShapeText [-InputObject] <Object> [-Text] <string> [-PassThru] [<CommonParameters>]
```

### Run
```powershell
Set-OfficePowerPointShapeText [-InputObject] <Object> -Run <Object[]> [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Accepts either a PowerPointTextBox object or a PowerPointShapeInfo record
returned by Find-OfficePowerPointShape or Get-OfficePowerPointShape. This is the direct
object-editing counterpart to the creation DSL: locate the text box in an existing deck, replace its
contents, then save or close the presentation.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Find-OfficePowerPointShape -Presentation $ppt -Text 'Draft' -Kind TextBox |
                Set-OfficePowerPointShapeText -Text 'Ready'
```

Accepts shape metadata returned by Find-OfficePowerPointShape or Get-OfficePowerPointShape.

### EXAMPLE 2
```powershell
PS> $ppt = Get-OfficePowerPoint -Path .\Release.pptx
Find-OfficePowerPointShape -Presentation $ppt -Text 'Status marker' -Kind TextBox |
    Set-OfficePowerPointShapeText -Text 'Status marker: Ready for launch'
$ppt | Close-OfficePowerPoint -Save
```

Searches the existing deck, edits the matched text box, and saves the presentation.

## PARAMETERS

### -InputObject
PowerPoint text box or shape-info record for a text box to update.

```yaml
Type: Object
Parameter Sets: Text, Run
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -PassThru
Emit the updated text box so additional OfficeIMO operations can continue.

```yaml
Type: SwitchParameter
Parameter Sets: Text, Run
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Run
Replacement rich text runs. Each run can be created with TextRun/PowerPointTextRun or provided as a hashtable/object.

```yaml
Type: Object[]
Parameter Sets: Run
Aliases: Runs
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Text
Replacement text. A null value clears the text box.

```yaml
Type: String
Parameter Sets: Text
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.Object`

## OUTPUTS

- `OfficeIMO.PowerPoint.PowerPointTextBox`

## RELATED LINKS

- None
