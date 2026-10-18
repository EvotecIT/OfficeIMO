---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeWordComboBox
## SYNOPSIS
Adds a combo box content control to the current paragraph.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficeWordComboBox [-Items] <string[]> [-DefaultValue <string>] [-Alias <string>] [-Tag <string>] [-Paragraph <WordParagraph>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a combo box content control to the current paragraph.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Add-OfficeWordParagraph { Add-OfficeWordComboBox -Items 'Red','Blue' -DefaultValue 'Blue' -Alias 'Color' }
```

Creates a combo box control with the selected default value.

## PARAMETERS

### -Alias
Optional alias for the control.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DefaultValue
Optional default value (must match one of the items).

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Items
Items to include in the combo box.

```yaml
Type: String[]
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Paragraph
Explicit paragraph to receive the control.

```yaml
Type: WordParagraph
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -PassThru
Emit the created control.

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

### -Tag
Optional tag for the control.

```yaml
Type: String
Parameter Sets: __AllParameterSets
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

- `OfficeIMO.Word.WordParagraph`

## OUTPUTS

- `OfficeIMO.Word.WordComboBox`

## RELATED LINKS

- None
