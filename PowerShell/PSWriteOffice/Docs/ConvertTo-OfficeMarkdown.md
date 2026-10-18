---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# ConvertTo-OfficeMarkdown
## SYNOPSIS
Converts objects into a Markdown table.

## SYNTAX
### __AllParameterSets
```powershell
ConvertTo-OfficeMarkdown [-InputObject <Object>] [-CollectionSeparator <string>] [-DictionaryEntrySeparator <string>] [-DictionaryKeyValueSeparator <string>] [-MaxCollectionItems <int>] [-MaxNestingDepth <int>] [-DisableAutoAlign] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Returns Markdown text by default; use -PassThru to emit a MarkdownDoc.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $markdown = $data | ConvertTo-OfficeMarkdown
```

Generates Markdown table text from the input objects.

### EXAMPLE 2
```powershell
PS> $doc = $data | ConvertTo-OfficeMarkdown -PassThru
$doc.P('Totals above'); $doc.ToMarkdown()
```

Builds a table and appends more content using the MarkdownDoc API.

### EXAMPLE 3
```powershell
PS> $markdown = $data | ConvertTo-OfficeMarkdown -DisableAutoAlign
```

Forces left-aligned columns instead of auto-aligned output.

## PARAMETERS

### -CollectionSeparator
Text used between items when a cell contains a collection.

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

### -DictionaryEntrySeparator
Text used between entries when a cell contains a dictionary.

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

### -DictionaryKeyValueSeparator
Text used between a dictionary key and value.

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

### -DisableAutoAlign
Disable automatic alignment heuristics for tables.

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

### -InputObject
Objects to convert into Markdown.

```yaml
Type: Object
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -MaxCollectionItems
Maximum number of items allowed in one nested collection or dictionary cell. Defaults to 1,048,575; increase explicitly for trusted larger values.

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

### -MaxNestingDepth
Maximum nesting depth allowed while normalizing one cell value. Defaults to 64; increase explicitly for trusted deeper values.

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

### -PassThru
Emit a Markdown document object instead of text.

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

- `System.Object`

## OUTPUTS

- `System.String`
- `OfficeIMO.Markdown.MarkdownDoc`

## RELATED LINKS

- None
