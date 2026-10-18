---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# ConvertFrom-OfficeRtf
## SYNOPSIS
Converts RTF input to Word, HTML, PDF, or Markdown output.

## SYNTAX
### Path (Default)
```powershell
ConvertFrom-OfficeRtf [-Path] <string> -As <OfficeRtfConversionTarget> [-OutputPath <string>] [-FontFamily <string>] [-IncludeFontStyles] [-IncludeListStyles] [-IncludeParagraphClasses] [-IncludeRunClasses] [-IncludeDefaultCss] [-UseImagePaths] [-IncludeHiddenText] [-ExcludeImages] [-ExcludeTables] [-ExcludeHeaderFooters] [-ExcludeNotes] [-NoUnsupportedHtmlComments] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Text
```powershell
ConvertFrom-OfficeRtf -Text <string> -As <OfficeRtfConversionTarget> [-OutputPath <string>] [-FontFamily <string>] [-IncludeFontStyles] [-IncludeListStyles] [-IncludeParagraphClasses] [-IncludeRunClasses] [-IncludeDefaultCss] [-UseImagePaths] [-IncludeHiddenText] [-ExcludeImages] [-ExcludeTables] [-ExcludeHeaderFooters] [-ExcludeNotes] [-NoUnsupportedHtmlComments] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Converts RTF input to Word, HTML, PDF, or Markdown output.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficeRtf -Path .\Report.rtf -Text 'Summary', 'Ready for review'
            ConvertFrom-OfficeRtf -Path .\Report.rtf -As Word -OutputPath .\Report.docx -PassThru
```

Loads the RTF file and saves a Word document using OfficeIMO.Word.Rtf.

### EXAMPLE 2
```powershell
PS> ConvertFrom-OfficeRtf -Path .\Report.rtf -As Html -OutputPath .\Report.html
```

Converts RTF to Word, then serializes Word to HTML.

### EXAMPLE 3
```powershell
PS> ConvertFrom-OfficeRtf -Path .\Report.rtf -As Pdf -OutputPath .\Report.pdf
```

Uses OfficeIMO.Rtf.Pdf to save a PDF file.

### EXAMPLE 4
```powershell
PS> ConvertFrom-OfficeRtf -Path .\Report.rtf -As Markdown -OutputPath .\Report.md -PassThru
```

Converts the RTF document to Markdown using OfficeIMO.Rtf.Markdown.

## PARAMETERS

### -As
Target document format.

```yaml
Type: OfficeRtfConversionTarget
Parameter Sets: Path, Text
Aliases: None
Possible values: Word, Html, Pdf, Markdown

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ExcludeHeaderFooters
Exclude RTF headers and footers from PDF output.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ExcludeImages
Exclude RTF images from PDF output.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ExcludeNotes
Exclude RTF notes from PDF output.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ExcludeTables
Exclude RTF tables from PDF output.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FontFamily
Optional font family for RTF to HTML conversion.

```yaml
Type: String
Parameter Sets: Path, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IncludeDefaultCss
Include the built-in default CSS for RTF to HTML conversion.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IncludeFontStyles
Include font styles as inline CSS for RTF to HTML conversion.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IncludeHiddenText
Include hidden RTF text when converting to PDF or Markdown.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IncludeListStyles
Include list style metadata for RTF to HTML conversion.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IncludeParagraphClasses
Emit paragraph styles as CSS classes for RTF to HTML conversion.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IncludeRunClasses
Emit run styles as CSS classes for RTF to HTML conversion.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoUnsupportedHtmlComments
Do not emit HTML comments for unsupported RTF features when converting to Markdown.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputPath
Optional output path. When omitted, the converted object or text is returned.

```yaml
Type: String
Parameter Sets: Path, Text
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
Parameter Sets: Path, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
RTF file path.

```yaml
Type: String
Parameter Sets: Path
Aliases: FilePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Text
Raw RTF text.

```yaml
Type: String
Parameter Sets: Text
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -UseImagePaths
Store image references as file paths instead of base64 data URIs for HTML output.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Text
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

- `System.String`

## OUTPUTS

- `OfficeIMO.Word.WordDocument`
- `System.String`
- `OfficeIMO.Pdf.PdfDocument`
- `System.IO.FileInfo`

## RELATED LINKS

- None
