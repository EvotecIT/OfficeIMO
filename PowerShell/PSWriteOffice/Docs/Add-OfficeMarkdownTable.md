---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeMarkdownTable
## SYNOPSIS
Adds a Markdown table from objects.

## SYNTAX
### Context (Default)
```powershell
Add-OfficeMarkdownTable [-InputObject] <Object> [-View <OfficeTableView>] [-CollectionSeparator <string>] [-DictionaryEntrySeparator <string>] [-DictionaryKeyValueSeparator <string>] [-MaxCollectionItems <int>] [-MaxNestingDepth <int>] [-DisableAutoAlign] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficeMarkdownTable [-InputObject] <Object> -Document <MarkdownDoc> [-View <OfficeTableView>] [-CollectionSeparator <string>] [-DictionaryEntrySeparator <string>] [-DictionaryKeyValueSeparator <string>] [-MaxCollectionItems <int>] [-MaxNestingDepth <int>] [-DisableAutoAlign] [-PassThru] [<CommonParameters>]
```

### PipelineDocument
```powershell
Add-OfficeMarkdownTable [-InputObject] <Object> -Document <MarkdownDoc> [-View <OfficeTableView>] [-CollectionSeparator <string>] [-DictionaryEntrySeparator <string>] [-DictionaryKeyValueSeparator <string>] [-MaxCollectionItems <int>] [-MaxNestingDepth <int>] [-DisableAutoAlign] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a Markdown table from objects.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> MarkdownTable -InputObject $rows
```

Appends a Markdown table using the supplied objects.

### EXAMPLE 2
```powershell
PS> $doc = New-OfficeMarkdown -Path .\Report.md -NoSave -PassThru
$doc | MarkdownTable -InputObject $summary -PassThru | MarkdownTable -InputObject $details
```

Creates two tables in sequence within the same Markdown document.

## PARAMETERS

### -CollectionSeparator
Text used between items when a cell contains a collection.

```yaml
Type: String
Parameter Sets: Context, Document, PipelineDocument
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
Parameter Sets: Context, Document, PipelineDocument
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
Parameter Sets: Context, Document, PipelineDocument
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
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Markdown document to update outside the DSL context.

```yaml
Type: MarkdownDoc
Parameter Sets: Document, PipelineDocument
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -InputObject
Objects to convert into a Markdown table.

```yaml
Type: Object
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -MaxCollectionItems
Maximum number of items allowed in one nested collection or dictionary cell. Defaults to 1,048,575; increase explicitly for trusted larger values.

```yaml
Type: Int32
Parameter Sets: Context, Document, PipelineDocument
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
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the Markdown document after appending the table.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -View
Projection to apply before writing the table.

```yaml
Type: OfficeTableView
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values: Normal, Transpose

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Markdown.MarkdownDoc`
- `System.Object`

## OUTPUTS

- `OfficeIMO.Markdown.MarkdownDoc`

## RELATED LINKS

- None
