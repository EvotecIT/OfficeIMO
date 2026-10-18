---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeWordField
## SYNOPSIS
Gets fields from a Word document.

## SYNTAX
### Path (Default)
```powershell
Get-OfficeWordField [-Path] <string> [-FieldType <WordFieldType[]>] [-Contains <string>] [-CaseSensitive] [<CommonParameters>]
```

### Document
```powershell
Get-OfficeWordField -Document <WordDocument> [-FieldType <WordFieldType[]>] [-Contains <string>] [-CaseSensitive] [<CommonParameters>]
```

## DESCRIPTION
Returns WordField objects, optionally filtered by type or field code.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $fields = Get-OfficeWordField -Path .\Report.docx -Contains 'TOC'
$fields |
    Select-Object -Property FieldType, Field |
    Format-Table -AutoSize
```

Filters Word fields by field-code text before running a field or table-of-contents refresh.

## PARAMETERS

### -CaseSensitive
Use case-sensitive matching for Contains.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Contains
Filter by field code text.

```yaml
Type: String
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Word document to read.

```yaml
Type: WordDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -FieldType
Filter by field type.

```yaml
Type: WordFieldType[]
Parameter Sets: Path, Document
Aliases: None
Possible values: AddressBlock, Advance, Ask, Author, AutoNum, AutoNumLgl, AutoNumOut, AutoText, AutoTextList, Bibliography, Citation, Comments, Compare, CreateDate, Database, Date, DocProperty, DocVariable, Embed, FileName, FileSize, GoToButton, GreetingLine, HyperlinkIf, IncludePicture, IncludeText, Index, Info, Keywords, LastSavedBy, Link, ListNum, MacroButton, MergeField, MergeRec, MergeSeq, Next, NextIf, NoteRef, NumChars, NumPages, NumWords, Page, PageRef, Print, PrintDate, Private, Quote, RD, Ref, RevNum, SaveDate, Section, SectionPages, Seq, Set, SkipIf, StyleRef, Subject, Symbol, TA, TC, Template, Time, Title, TOA, TOC, UserAddress, UserInitials, UserName, XE, Formula, EQ

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Path to the .docx file.

```yaml
Type: String
Parameter Sets: Path
Aliases: InputPath, FilePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Word.WordDocument`

## OUTPUTS

- `OfficeIMO.Word.WordField`

## RELATED LINKS

- None
