---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Update-OfficeWordText
## SYNOPSIS
Replaces text in a Word document.

## SYNTAX
### Auto (Default)
```powershell
Update-OfficeWordText -OldValue <string> -NewValue <string> [-CaseSensitive] [-IncludeHyperlinkText] [-IncludeHyperlinkUri] [-IncludeHyperlinkAnchor] [-IncludeHyperlinkTooltip] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
Update-OfficeWordText -OldValue <string> -NewValue <string> [-Document <WordDocument>] [-CaseSensitive] [-IncludeHyperlinkText] [-IncludeHyperlinkUri] [-IncludeHyperlinkAnchor] [-IncludeHyperlinkTooltip] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Path
```powershell
Update-OfficeWordText [-Path] <string> -OldValue <string> -NewValue <string> [-CaseSensitive] [-IncludeHyperlinkText] [-IncludeHyperlinkUri] [-IncludeHyperlinkAnchor] [-IncludeHyperlinkTooltip] [-Open] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Supports direct document objects, file paths, and the active DSL document. Hyperlink labels and metadata can be updated when requested.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $count = $doc | Update-OfficeWordText -OldValue 'FY24' -NewValue 'FY25' -PassThru
```

Updates matching text in the loaded document and returns the number of replacements.

### EXAMPLE 2
```powershell
PS> Update-OfficeWordText -Path .\Report.docx -OldValue 'old.example.com' -NewValue 'new.example.com' -IncludeHyperlinkUri
```

Loads the document, updates matching hyperlink URLs, saves the file, and closes it.

## PARAMETERS

### -CaseSensitive
Use case-sensitive matching.

```yaml
Type: SwitchParameter
Parameter Sets: Auto, Document, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Document to update.

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

### -IncludeHyperlinkAnchor
Also replace hyperlink anchors.

```yaml
Type: SwitchParameter
Parameter Sets: Auto, Document, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IncludeHyperlinkText
Also replace hyperlink display text.

```yaml
Type: SwitchParameter
Parameter Sets: Auto, Document, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IncludeHyperlinkTooltip
Also replace hyperlink tooltips.

```yaml
Type: SwitchParameter
Parameter Sets: Auto, Document, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IncludeHyperlinkUri
Also replace hyperlink URIs.

```yaml
Type: SwitchParameter
Parameter Sets: Auto, Document, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NewValue
Replacement text.

```yaml
Type: String
Parameter Sets: Auto, Document, Path
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OldValue
Text to find.

```yaml
Type: String
Parameter Sets: Auto, Document, Path
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Open
Open the file after saving when using -Path.

```yaml
Type: SwitchParameter
Parameter Sets: Path
Aliases: Show
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the object created or changed by the command.

```yaml
Type: SwitchParameter
Parameter Sets: Auto, Document, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Path to the .docx file to update in place.

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

- `System.Int32`

## RELATED LINKS

- None
