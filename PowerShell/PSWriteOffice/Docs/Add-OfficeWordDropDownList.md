---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeWordDropDownList
## SYNOPSIS
Adds a dropdown list content control to the current paragraph.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficeWordDropDownList [-Items] <string[]> [-Alias <string>] [-Tag <string>] [-Paragraph <WordParagraph>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a dropdown list content control to the current paragraph.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Add-OfficeWordParagraph { Add-OfficeWordDropDownList -Items 'Low','Medium','High' -Alias 'Priority' }
```

Creates a dropdown list control with the supplied items.

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

### -Items
Items to include in the dropdown list.

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

- `OfficeIMO.Word.WordDropDownList`

## RELATED LINKS

- None
