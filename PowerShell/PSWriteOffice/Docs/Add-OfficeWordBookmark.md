---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeWordBookmark
## SYNOPSIS
Adds a bookmark to the current paragraph.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficeWordBookmark [-Name] <string> [-Paragraph <WordParagraph>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a bookmark to the current paragraph.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Add-OfficeWordParagraph { Add-OfficeWordText -Text 'Intro'; Add-OfficeWordBookmark -Name 'Intro' }
```

Creates a bookmark named Intro on the paragraph.

## PARAMETERS

### -Name
Bookmark name.

```yaml
Type: String
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
Explicit paragraph to receive the bookmark.

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
Emit the paragraph after adding the bookmark.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Word.WordParagraph`

## OUTPUTS

- `None`

## RELATED LINKS

- None
