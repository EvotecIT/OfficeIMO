---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficeMarkdownPdfOptions
## SYNOPSIS
Creates native Markdown-to-PDF conversion options for Convert-OfficeDocument.

## SYNTAX
### Default (Default)
```powershell
New-OfficeMarkdownPdfOptions [-BaseOptions <MarkdownPdfExportOptions>] [-PdfOptions <PdfCreateOptions>] [-Theme <OfficeThemeInput>] [-Font <string>] [-Title <string>] [-Author <string>] [-Subject <string>] [-Keywords <string>] [-BaseDirectory <string>] [-NoDefaultTheme] [-IncludeLocalImages] [-IncludeDataUriImages] [-NoLocalImageRestriction] [-MaximumDataUriImageBytes <Int32>] [-DefaultImageWidth <OfficeLength>] [-DefaultImageHeight <OfficeLength>] [-FrontMatterRenderMode <MarkdownPdfFrontMatterRenderMode>] [-NoFrontMatterTheme] [-NoFrontMatterMetadata] [-NoFirstHeadingTitle] [-NoOutline] [<CommonParameters>]
```

## DESCRIPTION
Creates native Markdown-to-PDF conversion options for Convert-OfficeDocument.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $options = New-OfficeMarkdownPdfOptions -Title 'Service report' -Author 'Evotec' -IncludeLocalImages -BaseDirectory .\Assets
Convert-OfficeDocument -Path .\Report.md -OutputPath .\Report.pdf -Options $options
```

Builds a typed options object through ordinary PowerShell parameters; no hashtable or .NET construction is required.

## PARAMETERS

### -Author
PDF author metadata.

```yaml
Type: String
Parameter Sets: Default
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
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -BaseOptions
Existing Markdown PDF options to clone and override.

```yaml
Type: MarkdownPdfExportOptions
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -DefaultImageHeight
Fallback image height in PDF points.

```yaml
Type: OfficeLength
Parameter Sets: Default
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
Type: OfficeLength
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Font
Default font family.

```yaml
Type: String
Parameter Sets: Default
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
Parameter Sets: Default
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
Parameter Sets: Default
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
Parameter Sets: Default
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
Parameter Sets: Default
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
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoDefaultTheme
Skip the native default Markdown PDF theme.

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

### -NoFirstHeadingTitle
Use the first Markdown heading as the PDF title when no title is supplied.

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

### -NoFrontMatterMetadata
Use front matter values as PDF metadata.

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

### -NoFrontMatterTheme
Use front matter values to select a visual theme.

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

### -NoLocalImageRestriction
Require local images to resolve under BaseDirectory.

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

### -NoOutline
Create PDF outlines from Markdown headings.

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

### -PdfOptions
Underlying low-level OfficeIMO PDF options.

```yaml
Type: PdfCreateOptions
Parameter Sets: Default
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
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Theme
Built-in theme name, shared theme or theme package.

```yaml
Type: OfficeThemeInput
Parameter Sets: Default
Aliases: None
Possible values:

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

- `OfficeIMO.Markdown.Pdf.MarkdownPdfExportOptions`

## OUTPUTS

- `OfficeIMO.Markdown.Pdf.MarkdownPdfExportOptions`

## RELATED LINKS

- None
