---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficeWordField
## SYNOPSIS
Gets fields, filtered by type or field code text.

## SYNTAX
### Document (Default)
```powershell
Get-OfficeWordField [-FieldType <WordFieldType[]>] [-Contains <string>] [-CaseSensitive] [-Document <WordDocument>] [<CommonParameters>]
```

### Path
```powershell
Get-OfficeWordField [-Path] <string[]> [-FieldType <WordFieldType[]>] [-Contains <string>] [-CaseSensitive] [<CommonParameters>]
```

### LiteralPath
```powershell
Get-OfficeWordField -LiteralPath <string[]> [-FieldType <WordFieldType[]>] [-Contains <string>] [-CaseSensitive] [<CommonParameters>]
```

## DESCRIPTION
Gets fields, filtered by type or field code text.

## EXAMPLES

### EXAMPLE 1
```powershell
Get-OfficeWordField -CaseSensitive
```


### EXAMPLE 2
```powershell
Get-OfficeWordField -LiteralPath @('C:\Path')
```


### EXAMPLE 3
```powershell
Get-OfficeWordField -Path @('C:\Path')
```


## PARAMETERS

### -CaseSensitive
CaseSensitive for this inspection.

```yaml
Type: SwitchParameter
Parameter Sets: Document, Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Contains
Contains for this inspection.

```yaml
Type: String
Parameter Sets: Document, Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Document for this inspection.

```yaml
Type: WordDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -FieldType
FieldType for this inspection.

```yaml
Type: WordFieldType[]
Parameter Sets: Document, Path, LiteralPath
Aliases: None
Possible values: AddressBlock, Advance, Ask, Author, AutoNum, AutoNumLgl, AutoNumOut, AutoText, AutoTextList, Bibliography, Citation, Comments, Compare, CreateDate, Database, Date, DocProperty, DocVariable, Embed, FileName, FileSize, GoToButton, GreetingLine, HyperlinkIf, IncludePicture, IncludeText, Index, Info, Keywords, LastSavedBy, Link, ListNum, MacroButton, MergeField, MergeRec, MergeSeq, Next, NextIf, NoteRef, NumChars, NumPages, NumWords, Page, PageRef, Print, PrintDate, Private, Quote, RD, Ref, RevNum, SaveDate, Section, SectionPages, Seq, Set, SkipIf, StyleRef, Subject, Symbol, TA, TC, Template, Time, Title, TOA, TOC, UserAddress, UserInitials, UserName, XE, Formula, EQ

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LiteralPath
LiteralPath for this inspection.

```yaml
Type: String[]
Parameter Sets: LiteralPath
Aliases: PSPath, FullName
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -Path
Path for this inspection.

```yaml
Type: String[]
Parameter Sets: Path
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.String[]`
- `OfficeIMO.Word.WordDocument`

## OUTPUTS

- `OfficeIMO.Word.WordField`

## RELATED LINKS

- None
