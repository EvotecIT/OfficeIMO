---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# ConvertFrom-OfficeOneNote
## SYNOPSIS
Converts an offline OneNote section or notebook to semantic Markdown, HTML, or PDF.

## SYNTAX
### __AllParameterSets
```powershell
ConvertFrom-OfficeOneNote [-Path] <string> [-OutputPath] <string> [-ReadOptions <OneNoteReaderOptions>] [-NotebookOptions <OneNoteNotebookReaderOptions>] [-ProjectionOptions <OneNoteMarkdownOptions>] [-HtmlOptions <HtmlOptions>] [-PdfOptions <OneNoteToPdfOptions>] [-FailOnLoss] [-Force] [-PassThruReport] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Free-form canvas placement and unsupported native data are reported as conversion evidence rather than silently presented as lossless.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $report = ConvertFrom-OfficeOneNote -Path .\Operations.one -OutputPath .\Operations.md -PassThruReport
$report | Select-Object HasLoss, Diagnostics
```


## PARAMETERS

### -FailOnLoss
Fail when the selected projection reports an approximation or omission.

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

### -Force
Overwrite an existing destination.

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

### -HtmlOptions
HTML rendering settings used for .html output.

```yaml
Type: HtmlOptions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NotebookOptions
Notebook hierarchy, package, and section-error policy.

```yaml
Type: OneNoteNotebookReaderOptions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputPath
Destination .md, .html, or .pdf path.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: OutPath
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThruReport
Return conversion evidence instead of file information.

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

### -Path
Path to a .one section, .onetoc2 notebook index, or .onepkg archive.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: FilePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -PdfOptions
Semantic layout and PDF settings used for .pdf output.

```yaml
Type: OneNoteToPdfOptions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ProjectionOptions
OneNote hierarchy, history, and binary-asset projection settings.

```yaml
Type: OneNoteMarkdownOptions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ReadOptions
Bounded section and revision-store read options.

```yaml
Type: OneNoteReaderOptions
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

- `System.String`

## OUTPUTS

- `System.IO.FileInfo`
- `OfficeIMO.OneNote.Markdown.OneNoteMarkdownConversionReport`
- `OfficeIMO.Html.HtmlConversionReport`
- `OfficeIMO.Pdf.PdfDocumentConversionResult`

## RELATED LINKS

- None
