---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# ConvertTo-OfficeRtf
## SYNOPSIS
Converts Word, HTML, PDF, or Markdown input to RTF.

## SYNTAX
### WordPath (Default)
```powershell
ConvertTo-OfficeRtf -WordPath <string> [-OutputPath <string>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### WordDocument
```powershell
ConvertTo-OfficeRtf -WordDocument <WordDocument> [-OutputPath <string>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Html
```powershell
ConvertTo-OfficeRtf -Html <string> [-OutputPath <string>] [-FontFamily <string>] [-BasePath <string>] [-StylesheetPath <string[]>] [-StylesheetContent <string[]>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### HtmlPath
```powershell
ConvertTo-OfficeRtf -HtmlPath <string> [-OutputPath <string>] [-FontFamily <string>] [-BasePath <string>] [-StylesheetPath <string[]>] [-StylesheetContent <string[]>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### PdfPath
```powershell
ConvertTo-OfficeRtf -PdfPath <string> [-OutputPath <string>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Markdown
```powershell
ConvertTo-OfficeRtf -Markdown <string> [-OutputPath <string>] [-PreserveRawHtmlAsText] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### MarkdownPath
```powershell
ConvertTo-OfficeRtf -MarkdownPath <string> [-OutputPath <string>] [-PreserveRawHtmlAsText] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Converts Word, HTML, PDF, or Markdown input to RTF.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficeWord -Path .\Report.docx { WordParagraph -Text 'Summary' }
            ConvertTo-OfficeRtf -WordPath .\Report.docx -OutputPath .\Report.rtf -PassThru
```

Loads the Word document and saves an RTF file using OfficeIMO.Word.Rtf.

### EXAMPLE 2
```powershell
PS> ConvertTo-OfficeRtf -Html '<h1>Report</h1>' -OutputPath .\Report.rtf
```

Creates a Word document from HTML and serializes it to RTF.

### EXAMPLE 3
```powershell
PS> ConvertTo-OfficeRtf -PdfPath .\Report.pdf -OutputPath .\Report.rtf
```

Uses OfficeIMO.Rtf.Pdf's semantic PDF reader to write RTF output.

### EXAMPLE 4
```powershell
PS> ConvertTo-OfficeRtf -MarkdownPath .\Report.md -OutputPath .\Report.rtf -PassThru
```

Parses Markdown and writes RTF using OfficeIMO.Rtf.Markdown.

## PARAMETERS

### -BasePath
Base path used to resolve relative HTML resources.

```yaml
Type: String
Parameter Sets: Html, HtmlPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FontFamily
Optional font family for HTML to Word conversion before RTF serialization.

```yaml
Type: String
Parameter Sets: Html, HtmlPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Html
HTML markup to convert to RTF through the Word HTML converter.

```yaml
Type: String
Parameter Sets: Html
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HtmlPath
Path to an HTML file to convert to RTF.

```yaml
Type: String
Parameter Sets: HtmlPath
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Markdown
Markdown text to convert to RTF.

```yaml
Type: String
Parameter Sets: Markdown
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MarkdownPath
Path to a Markdown file to convert to RTF.

```yaml
Type: String
Parameter Sets: MarkdownPath
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputPath
Optional destination RTF path. When omitted, raw RTF text is returned.

```yaml
Type: String
Parameter Sets: WordPath, WordDocument, Html, HtmlPath, PdfPath, Markdown, MarkdownPath
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
Parameter Sets: WordPath, WordDocument, Html, HtmlPath, PdfPath, Markdown, MarkdownPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PdfPath
Path to a PDF file to convert to semantic RTF.

```yaml
Type: String
Parameter Sets: PdfPath
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PreserveRawHtmlAsText
Preserve raw HTML Markdown blocks as plain text in the RTF output.

```yaml
Type: SwitchParameter
Parameter Sets: Markdown, MarkdownPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -StylesheetContent
Inline CSS stylesheets to apply during HTML conversion.

```yaml
Type: String[]
Parameter Sets: Html, HtmlPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -StylesheetPath
Paths to CSS stylesheets to apply during HTML conversion.

```yaml
Type: String[]
Parameter Sets: Html, HtmlPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -WordDocument
Word document instance to convert to RTF.

```yaml
Type: WordDocument
Parameter Sets: WordDocument
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -WordPath
Path to a .docx file to convert to RTF.

```yaml
Type: String
Parameter Sets: WordPath
Aliases: None
Possible values:

Required: True
Position: named
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
