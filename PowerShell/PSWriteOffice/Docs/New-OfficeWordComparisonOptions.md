---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficeWordComparisonOptions
## SYNOPSIS
Creates discoverable structural comparison settings for Compare-OfficeDocument.

## SYNTAX
### Default (Default)
```powershell
New-OfficeWordComparisonOptions [-IgnoreWhitespace] [-IgnoreCase] [-NoRunFormatting] [-NoEffectiveFormatting] [-NoParagraphStyleIds] [-NoRunStyleIds] [-IncludeScope <WordComparisonScope[]>] [-ExcludeScope <WordComparisonScope[]>] [-NoFields] [-NoContentControls] [-NoBookmarks] [-NoHyperlinks] [-NoLists] [-NoComments] [-NoCommentAuthors] [-NoCommentText] [-NoCommentResolvedState] [-NoCommentTargets] [-NoCommentReplies] [-NoRevisions] [-NoRevisionAuthors] [-NoRevisionText] [-NoRevisionLocations] [-NoImages] [-NoShapes] [-NoBlockOrder] [-NoGeneratedIds] [-NoVolatileMetadata] [<CommonParameters>]
```

## DESCRIPTION
Creates discoverable structural comparison settings for Compare-OfficeDocument.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $options = New-OfficeWordComparisonOptions -IgnoreWhitespace -IgnoreCase -NoVolatileMetadata:$false
Compare-OfficeDocument -ReferencePath .\Before.docx -DifferencePath .\After.docx -Options $options
```


## PARAMETERS

### -ExcludeScope
Remove these comparison scopes from results.

```yaml
Type: WordComparisonScope[]
Parameter Sets: Default
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
Parameter Sets: Default
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
Parameter Sets: Default
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
Parameter Sets: Default
Aliases: None
Possible values: Paragraph, Run, Field, ContentControl, Bookmark, Hyperlink, List, Comment, Revision, Table, TableRow, TableCell, Image, Shape

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoBlockOrder
Compare document block order.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoBookmarks
Compare bookmarks.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoCommentAuthors
Compare comment authors.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoCommentReplies
Compare comment replies.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoCommentResolvedState
Compare comment resolved state.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoComments
Compare comments.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoCommentTargets
Compare comment targets.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoCommentText
Compare comment text.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoContentControls
Compare content controls.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoEffectiveFormatting
Compare resolved effective formatting.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoFields
Compare fields.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoGeneratedIds
Compare generated identifiers.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoHyperlinks
Compare hyperlinks.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoImages
Compare images.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoLists
Compare lists.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoParagraphStyleIds
Compare paragraph style identifiers.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoRevisionAuthors
Compare revision authors.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoRevisionLocations
Compare revision locations.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoRevisions
Compare tracked revisions.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoRevisionText
Compare revision text.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoRunFormatting
Compare direct run formatting.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoRunStyleIds
Compare run style identifiers.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoShapes
Compare supported shapes.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoVolatileMetadata
Compare volatile timestamps and metadata.

```yaml
Type: SwitchParameter
Parameter Sets: Default
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

- `None`

## OUTPUTS

- `OfficeIMO.Word.WordComparisonOptions`

## RELATED LINKS

- None
