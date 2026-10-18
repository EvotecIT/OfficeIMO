---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# New-OfficeWordComparisonOptions
## SYNOPSIS
Creates discoverable structural comparison settings for Compare-OfficeWordDocument.

## SYNTAX
### __AllParameterSets
```powershell
New-OfficeWordComparisonOptions [-IgnoreWhitespace] [-IgnoreCase] [-CompareRunFormatting] [-CompareEffectiveFormatting] [-CompareParagraphStyleIds] [-CompareRunStyleIds] [-IncludeScope <WordComparisonScope[]>] [-ExcludeScope <WordComparisonScope[]>] [-CompareFields] [-CompareContentControls] [-CompareBookmarks] [-CompareHyperlinks] [-CompareLists] [-CompareComments] [-CompareCommentAuthors] [-CompareCommentText] [-CompareCommentResolvedState] [-CompareCommentTargets] [-CompareCommentReplies] [-CompareRevisions] [-CompareRevisionAuthors] [-CompareRevisionText] [-CompareRevisionLocations] [-CompareImages] [-CompareShapes] [-CompareBlockOrder] [-CompareGeneratedIds] [-CompareVolatileMetadata] [<CommonParameters>]
```

## DESCRIPTION
Creates discoverable structural comparison settings for Compare-OfficeWordDocument.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $options = New-OfficeWordComparisonOptions -IgnoreWhitespace -IgnoreCase -CompareVolatileMetadata:$false
Compare-OfficeWordDocument -ReferencePath .\Before.docx -DifferencePath .\After.docx -Options $options
```


## PARAMETERS

### -CompareBlockOrder
Compare document block order.

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

### -CompareBookmarks
Compare bookmarks.

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

### -CompareCommentAuthors
Compare comment authors.

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

### -CompareCommentReplies
Compare comment replies.

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

### -CompareCommentResolvedState
Compare comment resolved state.

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

### -CompareComments
Compare comments.

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

### -CompareCommentTargets
Compare comment targets.

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

### -CompareCommentText
Compare comment text.

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

### -CompareContentControls
Compare content controls.

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

### -CompareEffectiveFormatting
Compare resolved effective formatting.

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

### -CompareFields
Compare fields.

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

### -CompareGeneratedIds
Compare generated identifiers.

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

### -CompareHyperlinks
Compare hyperlinks.

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

### -CompareImages
Compare images.

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

### -CompareLists
Compare lists.

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

### -CompareParagraphStyleIds
Compare paragraph style identifiers.

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

### -CompareRevisionAuthors
Compare revision authors.

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

### -CompareRevisionLocations
Compare revision locations.

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

### -CompareRevisions
Compare tracked revisions.

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

### -CompareRevisionText
Compare revision text.

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

### -CompareRunFormatting
Compare direct run formatting.

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

### -CompareRunStyleIds
Compare run style identifiers.

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

### -CompareShapes
Compare supported shapes.

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

### -CompareVolatileMetadata
Compare volatile timestamps and metadata.

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

### -ExcludeScope
Remove these comparison scopes from results.

```yaml
Type: WordComparisonScope[]
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Paragraph, Run, Field, ContentControl, Bookmark, Hyperlink, List, Comment, Revision, Table, TableRow, TableCell, Image, Shape

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IgnoreCase
Ignore character casing.

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

### -IgnoreWhitespace
Ignore differences caused only by whitespace runs.

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

### -IncludeScope
Limit results to these comparison scopes.

```yaml
Type: WordComparisonScope[]
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Paragraph, Run, Field, ContentControl, Bookmark, Hyperlink, List, Comment, Revision, Table, TableRow, TableCell, Image, Shape

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `OfficeIMO.Word.WordComparisonOptions`

## RELATED LINKS

- None
