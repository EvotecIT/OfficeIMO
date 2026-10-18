---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# ConvertTo-OfficeWordMarkdown
## SYNOPSIS
Converts a Word document to Markdown.

## SYNTAX
### Path (Default)
```powershell
ConvertTo-OfficeWordMarkdown [-Path] <string> [-OutputPath <string>] [-FontFamily <string>] [-EnableUnderline] [-EnableHighlight] [-ImageExportMode <ImageExportMode>] [-ImageDirectory <string>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
ConvertTo-OfficeWordMarkdown -Document <WordDocument> [-OutputPath <string>] [-FontFamily <string>] [-EnableUnderline] [-EnableHighlight] [-ImageExportMode <ImageExportMode>] [-ImageDirectory <string>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Returns Markdown text or saves it to a file when -OutputPath is specified.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $markdown = ConvertTo-OfficeWordMarkdown -Path .\report.docx
```

Loads the document and returns Markdown markup.

### EXAMPLE 2
```powershell
PS> ConvertTo-OfficeWordMarkdown -Path .\report.docx -OutputPath .\report.md -PassThru
```

Writes report.md and returns the file info.

## PARAMETERS

### -Document
Word document instance to convert.

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

### -EnableHighlight
Wrap highlighted text with Markdown highlight markers.

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

### -EnableUnderline
Wrap underlined text with HTML underline tags.

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

### -FontFamily
Optional font family that should be treated as inline code.

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

### -ImageDirectory
Directory used when exporting images as files.

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

### -ImageExportMode
Controls how images are emitted during Markdown conversion.

```yaml
Type: ImageExportMode
Parameter Sets: Path, Document
Aliases: None
Possible values: Base64, File

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputPath
Optional output path for the Markdown file.

```yaml
Type: String
Parameter Sets: Path, Document
Aliases: OutPath
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit a FileInfo when saving to disk.

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

### -Path
Path to a .docx file.

```yaml
Type: String
Parameter Sets: Path
Aliases: FilePath
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

- `System.String`
- `System.IO.FileInfo`

## RELATED LINKS

- None
