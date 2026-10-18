---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# New-OfficeWordRevisionFilter
## SYNOPSIS
Creates a discoverable Word revision filter for Resolve-OfficeWordRevision.

## SYNTAX
### __AllParameterSets
```powershell
New-OfficeWordRevisionFilter [-Author <string>] [-RevisionId <string>] [-RevisionType <WordReviewRevisionType>] [-DateFrom <DateTime>] [-DateTo <DateTime>] [-LocationKind <WordReviewLocationKind>] [-PartUri <string>] [-InTable] [-NotInTable] [-InContentControl] [-NotInContentControl] [-InTextBox] [-NotInTextBox] [<CommonParameters>]
```

## DESCRIPTION
Creates a discoverable Word revision filter for Resolve-OfficeWordRevision.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $filter = New-OfficeWordRevisionFilter -Author 'Alex' -InTable
Resolve-OfficeWordRevision -Path .\Review.docx -Action Accept -Filter $filter
```


## PARAMETERS

### -Author
Revision author.

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

### -DateFrom
Earliest revision date.

```yaml
Type: DateTime
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DateTo
Latest revision date.

```yaml
Type: DateTime
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -InContentControl
Limit results to revisions inside content controls.

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

### -InTable
Limit results to revisions inside tables.

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

### -InTextBox
Limit results to revisions inside text boxes.

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

### -LocationKind
Word part or container location kind.

```yaml
Type: WordReviewLocationKind
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Body, Header, Footer, Footnote, Endnote

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NotInContentControl
Limit results to revisions outside content controls.

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

### -NotInTable
Limit results to revisions outside tables.

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

### -NotInTextBox
Limit results to revisions outside text boxes.

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

### -PartUri
Package part URI.

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

### -RevisionId
Revision identifier.

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

### -RevisionType
Revision operation type.

```yaml
Type: WordReviewRevisionType
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Insertion, Deletion, MoveFrom, MoveTo, ParagraphFormatting, RunFormatting, TableFormatting, TableRowFormatting, TableCellFormatting, SectionFormatting, Unknown

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

- `OfficeIMO.Word.WordRevisionFilter`

## RELATED LINKS

- None
