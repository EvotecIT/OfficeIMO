---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficeWordField
## SYNOPSIS
Adds a field to the current paragraph.

## SYNTAX
### Type (Default)
```powershell
Add-OfficeWordField [-Type] <WordFieldType> [-Format <WordFieldFormat>] [-CustomFormat <string>] [-Advanced] [-Parameters <string[]>] [-Paragraph <WordParagraph>] [-PassThru] [<CommonParameters>]
```

### MergeField
```powershell
Add-OfficeWordField -MergeField <string> [-Format <WordFieldFormat>] [-CustomFormat <string>] [-Paragraph <WordParagraph>] [-PassThru] [<CommonParameters>]
```

### DocumentProperty
```powershell
Add-OfficeWordField -DocumentProperty <string> [-Format <WordFieldFormat>] [-CustomFormat <string>] [-Paragraph <WordParagraph>] [-PassThru] [<CommonParameters>]
```

### Bookmark
```powershell
Add-OfficeWordField -Bookmark <string> [-Format <WordFieldFormat>] [-CustomFormat <string>] [-Paragraph <WordParagraph>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Supports standard Word field codes such as Page, Date, or NumPages.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Add-OfficeWordParagraph { Add-OfficeWordField -Type Page }
```

Inserts a PAGE field into the paragraph.

## PARAMETERS

### -Advanced
Use advanced field representation.

```yaml
Type: SwitchParameter
Parameter Sets: Type
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Bookmark
Bookmark to reference.

```yaml
Type: String
Parameter Sets: Bookmark
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -CustomFormat
Custom format string (date/time fields).

```yaml
Type: String
Parameter Sets: Type, MergeField, DocumentProperty, Bookmark
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DocumentProperty
Built-in or custom document property to display.

```yaml
Type: String
Parameter Sets: DocumentProperty
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Format
Optional field format switch.

```yaml
Type: WordFieldFormat
Parameter Sets: Type, MergeField, DocumentProperty, Bookmark
Aliases: None
Possible values: Lower, Upper, FirstCap, Caps, MergeFormat, UpperRoman, LowerRoman, Arabic, LowerLetter, UpperLetter, Ordinal, OrdText, CardText, DollarText, Hex, CharFormat

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MergeField
Mail-merge field name.

```yaml
Type: String
Parameter Sets: MergeField
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Paragraph
Explicit paragraph to receive the field.

```yaml
Type: WordParagraph
Parameter Sets: Type, MergeField, DocumentProperty, Bookmark
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Parameters
Additional field parameters.

```yaml
Type: String[]
Parameter Sets: Type
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the paragraph after adding the field.

```yaml
Type: SwitchParameter
Parameter Sets: Type, MergeField, DocumentProperty, Bookmark
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Type
Field type to insert.

```yaml
Type: WordFieldType
Parameter Sets: Type
Aliases: None
Possible values: AddressBlock, Advance, Ask, Author, AutoNum, AutoNumLgl, AutoNumOut, AutoText, AutoTextList, Bibliography, Citation, Comments, Compare, CreateDate, Database, Date, DocProperty, DocVariable, Embed, FileName, FileSize, GoToButton, GreetingLine, HyperlinkIf, IncludePicture, IncludeText, Index, Info, Keywords, LastSavedBy, Link, ListNum, MacroButton, MergeField, MergeRec, MergeSeq, Next, NextIf, NoteRef, NumChars, NumPages, NumWords, Page, PageRef, Print, PrintDate, Private, Quote, RD, Ref, RevNum, SaveDate, Section, SectionPages, Seq, Set, SkipIf, StyleRef, Subject, Symbol, TA, TC, Template, Time, Title, TOA, TOC, UserAddress, UserInitials, UserName, XE, Formula, EQ

Required: True
Position: 0
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
