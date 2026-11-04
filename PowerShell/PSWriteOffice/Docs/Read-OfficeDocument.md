---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Read-OfficeDocument
## SYNOPSIS
Reads supported documents once and projects text, structure or ingestion views.

## SYNTAX
### Path (Default)
```powershell
Read-OfficeDocument [-Path] <string[]> [[-As] <OfficeReadView>] [-ByPage] [-NoPageMarkers] [-Bundle] [-Indented] [-NoChunks] [-HierarchyOptions <ReaderHierarchicalChunkingOptions>] [-StructureOptions <OfficeDocumentStructuredExtractionOptions>] [-Sheet <OfficeSheetSelector>] [-Range <string>] [-ChunkRows <Int32>] [-NoHeaderRow] [-NoFootnotes] [-NoNotes] [-NoHeadingChunks] [-NoHashes] [-IncludePageLocations] [-MaximumChunkCharacters <Int32>] [-MaximumTableRows <Int32>] [-MaximumInputBytes <Int64>] [-MaximumXmlCharactersInPart <Int64>] [-MaximumStoreItems <int>] [-Ocr] [-OcrOptions <OfficeOcrOptions>] [-Reader <OfficeDocumentReader>] [-Password <securestring>] [-Options <Object[]>] [-Recurse] [-Extension <string[]>] [-MaximumFiles <int>] [-MaximumTotalBytes <Int64>] [-ThrottleLimit <int>] [<CommonParameters>]
```

### InputObject
```powershell
Read-OfficeDocument [-InputObject] <OfficeDocumentReadResult> [[-As] <OfficeReadView>] [-ByPage] [-NoPageMarkers] [-Bundle] [-Indented] [-NoChunks] [-HierarchyOptions <ReaderHierarchicalChunkingOptions>] [-StructureOptions <OfficeDocumentStructuredExtractionOptions>] [<CommonParameters>]
```

### LiteralPath
```powershell
Read-OfficeDocument [[-As] <OfficeReadView>] -LiteralPath <string[]> [-ByPage] [-NoPageMarkers] [-Bundle] [-Indented] [-NoChunks] [-HierarchyOptions <ReaderHierarchicalChunkingOptions>] [-StructureOptions <OfficeDocumentStructuredExtractionOptions>] [-Sheet <OfficeSheetSelector>] [-Range <string>] [-ChunkRows <Int32>] [-NoHeaderRow] [-NoFootnotes] [-NoNotes] [-NoHeadingChunks] [-NoHashes] [-IncludePageLocations] [-MaximumChunkCharacters <Int32>] [-MaximumTableRows <Int32>] [-MaximumInputBytes <Int64>] [-MaximumXmlCharactersInPart <Int64>] [-MaximumStoreItems <int>] [-Ocr] [-OcrOptions <OfficeOcrOptions>] [-Reader <OfficeDocumentReader>] [-Password <securestring>] [-Options <Object[]>] [-Recurse] [-Extension <string[]>] [-MaximumFiles <int>] [-MaximumTotalBytes <Int64>] [-ThrottleLimit <int>] [<CommonParameters>]
```

## DESCRIPTION
Reads supported documents once and projects text, structure or ingestion views.

## EXAMPLES

### EXAMPLE 1
```powershell
Read-OfficeDocument ./Reports -Recurse -As Text
```


### EXAMPLE 2
```powershell
$document = Read-OfficeDocument ./Report.docx
$document | Read-OfficeDocument -As Tables
```


## PARAMETERS

### -As
Representation to return. Document is the default.

```yaml
Type: OfficeReadView
Parameter Sets: Path, InputObject, LiteralPath
Aliases: None
Possible values: Document, Text, Markdown, Json, Chunks, Hierarchy, Structure, Tables, Visuals, Assets, Ingest, Ocr

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Bundle
Return in-memory table or visual export bundles.

```yaml
Type: SwitchParameter
Parameter Sets: Path, InputObject, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ByPage
Return one text or Markdown object for each observed page, slide or worksheet.

```yaml
Type: SwitchParameter
Parameter Sets: Path, InputObject, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ChunkRows
Spreadsheet source rows per chunk.

```yaml
Type: Int32
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Extension
Selected source extensions, with or without a leading dot.

```yaml
Type: String[]
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HierarchyOptions
Hierarchical chunking policy.

```yaml
Type: ReaderHierarchicalChunkingOptions
Parameter Sets: Path, InputObject, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IncludePageLocations
Compute page locations where the native format does not store them.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Indented
Indent JSON output.

```yaml
Type: SwitchParameter
Parameter Sets: Path, InputObject, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -InputObject
Already read document; its source is never reopened.

```yaml
Type: OfficeDocumentReadResult
Parameter Sets: InputObject
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -LiteralPath
Gets or sets input paths without wildcard expansion, binding FileInfo.FullName and provider PSPath.

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

### -MaximumChunkCharacters
Maximum characters per source chunk.

```yaml
Type: Int32
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumFiles
Maximum selected files; zero means no limit.

```yaml
Type: Int32
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumInputBytes
Maximum bytes per source. Native owner limits still apply.

```yaml
Type: Int64
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumStoreItems
Maximum items per mail store; zero means all.

```yaml
Type: Int32
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumTableRows
Maximum rows per table chunk.

```yaml
Type: Int32
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumTotalBytes
Maximum combined source bytes accepted in this read batch.

```yaml
Type: Int64
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumXmlCharactersInPart
Maximum XML characters in one Office package part.

```yaml
Type: Int64
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoChunks
Omit chunks from an ingestion summary.

```yaml
Type: SwitchParameter
Parameter Sets: Path, InputObject, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoFootnotes
Exclude Word footnotes.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoHashes
Do not compute source and chunk hashes.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoHeaderRow
Treat the first spreadsheet row as data.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoHeadingChunks
Do not split Markdown at headings.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoNotes
Exclude PowerPoint speaker notes.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoPageMarkers
Omit portable page markers from Markdown.

```yaml
Type: SwitchParameter
Parameter Sets: Path, InputObject, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Ocr
Recognize text in source images and scanned pages using an optional OCR engine.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OcrOptions
Optional OCR provider and native recognition policies. Implies Ocr.

```yaml
Type: OfficeOcrOptions
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Options
Typed common and format read options, or common/format-keyed hashtables.

```yaml
Type: Object[]
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Password
Open password for encrypted sources.

```yaml
Type: SecureString
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Gets or sets input paths with provider wildcard expansion.

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

### -Range
A1 range within each selected worksheet.

```yaml
Type: String
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Reader
Caller-configured immutable Reader.

```yaml
Type: OfficeDocumentReader
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Recurse
Include subfolders.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Sheet
Spreadsheet worksheet name, one-based position or sheet object.

```yaml
Type: OfficeSheetSelector
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -StructureOptions
Structured extraction policy.

```yaml
Type: OfficeDocumentStructuredExtractionOptions
Parameter Sets: Path, InputObject, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ThrottleLimit
Gets or sets the maximum number of concurrent background operations.

```yaml
Type: Int32
Parameter Sets: Path, LiteralPath
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

- `OfficeIMO.Reader.OfficeDocumentReadResult`
- `System.String[]`

## OUTPUTS

- `OfficeIMO.Reader.OfficeDocumentReadResult`
- `System.String`
- `OfficeIMO.Reader.OfficeDocumentPageText`
- `OfficeIMO.Reader.OfficeDocumentPageMarkdown`
- `OfficeIMO.Reader.ReaderChunk`
- `OfficeIMO.Reader.ReaderChunkHierarchyResult`
- `OfficeIMO.Reader.OfficeDocumentStructuredExtractionResult`
- `OfficeIMO.Reader.ReaderTable`
- `OfficeIMO.Reader.ReaderTableExportBundle`
- `OfficeIMO.Reader.ReaderVisual`
- `OfficeIMO.Reader.ReaderVisualExportBundle`
- `OfficeIMO.Reader.OfficeDocumentAsset`
- `OfficeIMO.Reader.ReaderIngestResult`
- `OfficeIMO.Pdf.Ocr.PdfSearchableOcrReview`

## RELATED LINKS

- None
