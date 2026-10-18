---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeWordListItem
## SYNOPSIS
Adds a single list item.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficeWordListItem [[-Text] <string>] [-List <WordList>] [-Level <int>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Can be called within Add-OfficeWordList/WordList or against an existing WordList from the pipeline; supports nesting via -Level.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> WordList { Add-OfficeWordListItem -Text 'First task' }
```

Creates a bullet with the text “First task”.

### EXAMPLE 2
```powershell
PS> Find-OfficeWordList -Document $doc -Text 'Initial review' | Add-OfficeWordListItem -Text 'Final approval'
```

Finds a list in an existing document and appends a new item.

## PARAMETERS

### -Level
Zero-based list level.

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

### -List
Existing list to append to. When omitted, the current DSL list is used.

```yaml
Type: WordList
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
Emit the created WordParagraph.

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

### -Text
List item text.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Word.WordList`

## OUTPUTS

- `None`

## RELATED LINKS

- None
