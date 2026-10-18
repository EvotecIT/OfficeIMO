---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# New-OfficeMarkdownPdfOptions
## SYNOPSIS
Creates discoverable Markdown-to-PDF conversion options for Export-OfficeDocumentPdf.

## SYNTAX
### __AllParameterSets
```powershell
New-OfficeMarkdownPdfOptions [-Options <MarkdownToPdfOptions>] [-PdfOptions <PdfOptions>] [-Theme <OfficeVisualThemeKind>] [-FontFamily <string>] [-Title <string>] [-Author <string>] [-Subject <string>] [-Keywords <string>] [-BaseDirectory <string>] [-ApplyWordLikeTheme] [-IncludeLocalImages] [-IncludeDataUriImages] [-RestrictLocalImagesToBaseDirectory] [-MaximumDataUriImageBytes <Int32>] [-DefaultImageWidth <Double>] [-DefaultImageHeight <Double>] [-FrontMatterRenderMode <MarkdownPdfFrontMatterRenderMode>] [-UseFrontMatterVisualTheme] [-UseFrontMatterMetadata] [-UseFirstHeadingAsTitle] [-CreateOutlineFromHeadings] [<CommonParameters>]
```

## DESCRIPTION
Creates discoverable Markdown-to-PDF conversion options for Export-OfficeDocumentPdf.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $options = New-OfficeMarkdownPdfOptions -Title 'Service report' -Author 'Evotec' -IncludeLocalImages -BaseDirectory .\Assets
Export-OfficeDocumentPdf -InputPath .\Report.md -Path .\Report.pdf -MarkdownOptions $options
```

Builds a typed options object through ordinary PowerShell parameters; no hashtable or .NET construction is required.

## PARAMETERS

### -ApplyWordLikeTheme
Apply the built-in Word-like Markdown PDF baseline theme.

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

### -Author
PDF author metadata.

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

### -BaseDirectory
Base directory used to resolve local Markdown images.

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

### -CreateOutlineFromHeadings
Create PDF outlines from Markdown headings.

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

### -DefaultImageHeight
Fallback image height in PDF points.

```yaml
Type: Double
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DefaultImageWidth
Fallback image width in PDF points.

```yaml
Type: Double
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FontFamily
Default font family.

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

### -FrontMatterRenderMode
Controls how YAML front matter appears in the PDF body.

```yaml
Type: MarkdownPdfFrontMatterRenderMode
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Hidden, DocumentHeader, Table

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IncludeDataUriImages
Embed supported data URI images.

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

### -IncludeLocalImages
Embed supported local image files.

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

### -Keywords
PDF keywords metadata.

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

### -MaximumDataUriImageBytes
Maximum decoded bytes for one data URI image.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Options
Existing Markdown PDF options to clone and override.

```yaml
Type: MarkdownToPdfOptions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -PdfOptions
Underlying low-level OfficeIMO PDF options.

```yaml
Type: PdfOptions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RestrictLocalImagesToBaseDirectory
Require local images to resolve under BaseDirectory.

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

### -Subject
PDF subject metadata.

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

### -Theme
Built-in visual theme.

```yaml
Type: OfficeVisualThemeKind
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Plain, WordLike, TechnicalDocument, GitHubLike, Compact, Report

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Title
PDF title metadata.

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

### -UseFirstHeadingAsTitle
Use the first Markdown heading as the PDF title when no title is supplied.

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

### -UseFrontMatterMetadata
Use front matter values as PDF metadata.

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

### -UseFrontMatterVisualTheme
Use front matter values to select a visual theme.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Markdown.Pdf.MarkdownToPdfOptions`

## OUTPUTS

- `OfficeIMO.Markdown.Pdf.MarkdownToPdfOptions`

## RELATED LINKS

- None
