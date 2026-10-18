---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Export-OfficeDocumentPdf
## SYNOPSIS
Exports Word, Excel, PowerPoint, HTML, Markdown, RTF, and literal text documents to PDF.

## SYNTAX
### Document (Default)
```powershell
Export-OfficeDocumentPdf [-Document] <Object> [-Path] <string> [-Password <string>] [-WordOptions <WordToPdfOptions>] [-ExcelOptions <ExcelToPdfOptions>] [-PowerPointOptions <PowerPointToPdfOptions>] [-MarkdownOptions <MarkdownToPdfOptions>] [-RtfOptions <RtfToPdfOptions>] [-HtmlOptions <HtmlToPdfOptions>] [-TextOptions <PdfPlainTextOptions>] [-AllowLegacyImportLoss] [-MaximumInputBytes <long>] [-SourceConversionReportVariable <string>] [-PdfWarningVariable <string>] [-PdfConversionReportVariable <string>] [-Open] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Directory
```powershell
Export-OfficeDocumentPdf -InputDirectory <string> -OutputDirectory <string> [-CheckpointDirectory <string>] [-ConversionRouteId <string>] [-MaximumConcurrency <int>] [-MaximumFiles <int>] [-MaximumOutputBytes <long>] [-SourceExtensions <string[]>] [-NoRecurse] [-ConflictPolicy <string>] [-RetryFailed] [-ItemResults] [-SummaryVariable <string>] [-Password <string>] [-WordOptions <WordToPdfOptions>] [-ExcelOptions <ExcelToPdfOptions>] [-PowerPointOptions <PowerPointToPdfOptions>] [-MarkdownOptions <MarkdownToPdfOptions>] [-RtfOptions <RtfToPdfOptions>] [-HtmlOptions <HtmlToPdfOptions>] [-TextOptions <PdfPlainTextOptions>] [-AllowLegacyImportLoss] [-MaximumInputBytes <long>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Files
```powershell
Export-OfficeDocumentPdf -InputPaths <Object[]> -OutputDirectory <string> [-CheckpointDirectory <string>] [-ConversionRouteId <string>] [-MaximumConcurrency <int>] [-MaximumFiles <int>] [-MaximumOutputBytes <long>] [-ConflictPolicy <string>] [-RetryFailed] [-ItemResults] [-SummaryVariable <string>] [-Password <string>] [-WordOptions <WordToPdfOptions>] [-ExcelOptions <ExcelToPdfOptions>] [-PowerPointOptions <PowerPointToPdfOptions>] [-MarkdownOptions <MarkdownToPdfOptions>] [-RtfOptions <RtfToPdfOptions>] [-HtmlOptions <HtmlToPdfOptions>] [-TextOptions <PdfPlainTextOptions>] [-AllowLegacyImportLoss] [-MaximumInputBytes <long>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Path
```powershell
Export-OfficeDocumentPdf [-InputPath] <string> [-Path] <string> [-Password <string>] [-WordOptions <WordToPdfOptions>] [-ExcelOptions <ExcelToPdfOptions>] [-PowerPointOptions <PowerPointToPdfOptions>] [-MarkdownOptions <MarkdownToPdfOptions>] [-RtfOptions <RtfToPdfOptions>] [-HtmlOptions <HtmlToPdfOptions>] [-TextOptions <PdfPlainTextOptions>] [-AllowLegacyImportLoss] [-MaximumInputBytes <long>] [-SourceConversionReportVariable <string>] [-PdfWarningVariable <string>] [-PdfConversionReportVariable <string>] [-Open] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Accepts live OfficeIMO documents, individual files, directories, and selected file pipelines. Directory and file batches require PowerShell 7.4 or newer.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $document | Export-OfficeDocumentPdf -Path .\Report.pdf
```


### EXAMPLE 2
```powershell
PS> Export-OfficeDocumentPdf -InputPath .\Report.docx -Path .\Report.pdf -PassThru
```


### EXAMPLE 3
```powershell
PS> $options = New-OfficeMarkdownPdfOptions -Title 'Service report' -IncludeLocalImages -BaseDirectory .\Assets
Export-OfficeDocumentPdf -InputPath .\Report.md -Path .\Report.pdf -MarkdownOptions $options
```

Typed conversion options also apply to the matching formats in a mixed batch.

### EXAMPLE 4
```powershell
PS> Export-OfficeDocumentPdf -InputDirectory .\Documents -OutputDirectory .\PDF -CheckpointDirectory .\PDF-State
```


### EXAMPLE 5
```powershell
PS> Get-ChildItem .\Documents -File | Export-OfficeDocumentPdf -OutputDirectory .\PDF -ItemResults -SummaryVariable summary
```


## PARAMETERS

### -AllowLegacyImportLoss
Accept reported legacy DOC import loss. Known loss otherwise blocks conversion.

```yaml
Type: SwitchParameter
Parameter Sets: Document, Directory, Files, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -CheckpointDirectory
Optional durable state directory. Completed sources and outputs are hash-verified on restart.

```yaml
Type: String
Parameter Sets: Directory, Files
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ConflictPolicy
Destination conflict policy. Durable checkpoint jobs require Fail.

```yaml
Type: String
Parameter Sets: Directory, Files
Aliases: None
Possible values: Fail, Rename, Replace

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ConversionRouteId
Explicit executable conversion route, for example html-pdf for HTML stored in TXT files. By default the catalog selects each file's PDF route.

```yaml
Type: String
Parameter Sets: Directory, Files
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Live Word, Excel, PowerPoint, HTML, Markdown, or RTF document to export. Saved FileInfo and path strings from the pipeline are opened automatically.

```yaml
Type: Object
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -ExcelOptions
Excel-specific PDF options.

```yaml
Type: ExcelToPdfOptions
Parameter Sets: Document, Directory, Files, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HtmlOptions
HTML-specific PDF rendering options.

```yaml
Type: HtmlToPdfOptions
Parameter Sets: Document, Directory, Files, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -InputDirectory
Source directory for mixed-format PDF export. Batch execution requires PowerShell 7.4 or newer.

```yaml
Type: String
Parameter Sets: Directory
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -InputPath
Source .doc, .docx, .txt, .xlsx, .pptx, .html, .htm, .xhtml, .md, .markdown, or .rtf file.

```yaml
Type: String
Parameter Sets: Path
Aliases: SourcePath, FullName
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -InputPaths
Explicit source paths or FileInfo objects from Get-ChildItem.

```yaml
Type: Object[]
Parameter Sets: Files
Aliases: SourcePaths
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue, ByPropertyName)
Accept wildcard characters: False
```

### -ItemResults
Stream every structured per-file outcome instead of the final summary. Summary mode reports at most 100 detailed failure warnings and one suppressed-count warning.

```yaml
Type: SwitchParameter
Parameter Sets: Directory, Files
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MarkdownOptions
Markdown-specific PDF options.

```yaml
Type: MarkdownToPdfOptions
Parameter Sets: Document, Directory, Files, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumConcurrency
Concurrent conversions. Default two; bounds keep the batch resource budget explicit.

```yaml
Type: Int32
Parameter Sets: Directory, Files
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumFiles
Maximum discovered or explicitly selected source files.

```yaml
Type: Int32
Parameter Sets: Directory, Files
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumInputBytes
Maximum source bytes for DOC and TXT import.

```yaml
Type: Int64
Parameter Sets: Document, Directory, Files, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumOutputBytes
Maximum output bytes per document.

```yaml
Type: Int64
Parameter Sets: Directory, Files
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoRecurse
Exclude descendants when selecting a directory.

```yaml
Type: SwitchParameter
Parameter Sets: Directory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Open
Open the PDF after exporting it.

```yaml
Type: SwitchParameter
Parameter Sets: Document, Path
Aliases: Show
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputDirectory
PDF destination directory. Source names retain their extension plus .pdf.

```yaml
Type: String
Parameter Sets: Directory, Files
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the saved PDF file.

```yaml
Type: SwitchParameter
Parameter Sets: Document, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Password
Password used to open an encrypted Word, Excel, or PowerPoint source file.

```yaml
Type: String
Parameter Sets: Document, Directory, Files, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Destination PDF path.

```yaml
Type: String
Parameter Sets: Document, Path
Aliases: OutputPath, FilePath
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PdfConversionReportVariable
Variable name that receives the structured PDF conversion report.

```yaml
Type: String
Parameter Sets: Document, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PdfWarningVariable
Variable name that receives structured PDF conversion warnings.

```yaml
Type: String
Parameter Sets: Document, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PowerPointOptions
PowerPoint-specific PDF options.

```yaml
Type: PowerPointToPdfOptions
Parameter Sets: Document, Directory, Files, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RetryFailed
Retry recorded failures; completed items remain protected.

```yaml
Type: SwitchParameter
Parameter Sets: Directory, Files
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RtfOptions
RTF-specific PDF options.

```yaml
Type: RtfToPdfOptions
Parameter Sets: Document, Directory, Files, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SourceConversionReportVariable
Variable receiving source-stage reports separately from PDF rendering diagnostics.

```yaml
Type: String
Parameter Sets: Document, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SourceExtensions
Restrict directory discovery to these source extensions. Other files are reported as skipped.

```yaml
Type: String[]
Parameter Sets: Directory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SummaryVariable
Variable receiving the bounded batch summary, including when ItemResults is selected.

```yaml
Type: String
Parameter Sets: Directory, Files
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TextOptions
Literal text layout and strict decoding options. Applies only to TXT sources.

```yaml
Type: PdfPlainTextOptions
Parameter Sets: Document, Directory, Files, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -WordOptions
Word-specific PDF options.

```yaml
Type: WordToPdfOptions
Parameter Sets: Document, Directory, Files, Path
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

- `System.Object[]`
- `System.Object`
- `System.String`

## OUTPUTS

- `System.IO.FileInfo`
- `OfficeIMO.Workflows.OfficeConversionBatchResult`
- `OfficeIMO.Workflows.OfficeConversionBatchItemResult`

## RELATED LINKS

- None
