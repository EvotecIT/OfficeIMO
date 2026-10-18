---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# ConvertFrom-OfficeWordHtml
## SYNOPSIS
Creates a Word document from HTML.

## SYNTAX
### Html (Default)
```powershell
ConvertFrom-OfficeWordHtml [-Html] <string> [-OutputPath <string>] [-FontFamily <string>] [-BasePath <string>] [-StylesheetPath <string[]>] [-StylesheetContent <string[]>] [-IncludeListStyles] [-ContinueNumbering] [-SupportsHeadingNumbering] [-RenderPreAsTable] [-TableCaptionPosition <TableCaptionPosition>] [-SectionTagHandling <SectionTagHandling>] [-Open] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Path
```powershell
ConvertFrom-OfficeWordHtml [-Path] <string> [-OutputPath <string>] [-FontFamily <string>] [-BasePath <string>] [-StylesheetPath <string[]>] [-StylesheetContent <string[]>] [-IncludeListStyles] [-ContinueNumbering] [-SupportsHeadingNumbering] [-RenderPreAsTable] [-TableCaptionPosition <TableCaptionPosition>] [-SectionTagHandling <SectionTagHandling>] [-Open] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Returns a WordDocument or saves it to disk when -OutputPath is provided.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> ConvertFrom-OfficeWordHtml -Html '<h1>Hello</h1>' -OutputPath .\hello.docx
```

Writes a Word document containing the supplied HTML.

### EXAMPLE 2
```powershell
PS> $doc = ConvertFrom-OfficeWordHtml -Path .\snippet.html
```

Returns a Word document instance for further edits.

## PARAMETERS

### -BasePath
Base path used to resolve relative resources (for example images).

```yaml
Type: String
Parameter Sets: Html, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ContinueNumbering
Continue numbering across separate ordered lists.

```yaml
Type: SwitchParameter
Parameter Sets: Html, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FontFamily
Optional font family to apply during conversion.

```yaml
Type: String
Parameter Sets: Html, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Html
HTML markup to convert.

```yaml
Type: String
Parameter Sets: Html
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -IncludeListStyles
Include list style metadata.

```yaml
Type: SwitchParameter
Parameter Sets: Html, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Open
Open the document after saving.

```yaml
Type: SwitchParameter
Parameter Sets: Html, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputPath
Optional output path for the .docx file.

```yaml
Type: String
Parameter Sets: Html, Path
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
Parameter Sets: Html, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Path to an HTML file.

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

### -RenderPreAsTable
Render <pre> elements as single-cell tables.

```yaml
Type: SwitchParameter
Parameter Sets: Html, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SectionTagHandling
Controls how <section> tags are mapped into Word.

```yaml
Type: SectionTagHandling
Parameter Sets: Html, Path
Aliases: None
Possible values: WordSection, Block

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -StylesheetContent
Inline CSS stylesheets to apply during conversion.

```yaml
Type: String[]
Parameter Sets: Html, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -StylesheetPath
Paths to CSS stylesheets to apply during conversion.

```yaml
Type: String[]
Parameter Sets: Html, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SupportsHeadingNumbering
Convert headings into a numbered list.

```yaml
Type: SwitchParameter
Parameter Sets: Html, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TableCaptionPosition
Controls where table captions are emitted.

```yaml
Type: TableCaptionPosition
Parameter Sets: Html, Path
Aliases: None
Possible values: Above, Below

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
- `System.IO.FileInfo`

## RELATED LINKS

- None
