using OfficeIMO.Rtf.Diagnostics;

namespace OfficeIMO.Reader.Rtf;

/// <summary>
/// RTF ingestion adapter for <see cref="OfficeDocumentReader"/>.
/// </summary>
internal static partial class RtfReaderAdapter {
    /// <summary>
    /// Reads an RTF file and emits normalized chunks.
    /// </summary>
    public static IEnumerable<ReaderChunk> Read(string rtfPath, ReaderOptions? readerOptions = null, ReaderRtfOptions? rtfOptions = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        return ReadResult(rtfPath, readerOptions, rtfOptions, encoding, cancellationToken).Value;
    }

    /// <summary>Reads an RTF file and returns normalized chunks with operation-scoped fidelity diagnostics.</summary>
    public static RtfConversionResult<IReadOnlyList<ReaderChunk>> ReadResult(string rtfPath, ReaderOptions? readerOptions = null, ReaderRtfOptions? rtfOptions = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        if (rtfPath == null) throw new ArgumentNullException(nameof(rtfPath));
        if (rtfPath.Length == 0) throw new ArgumentException("RTF path cannot be empty.", nameof(rtfPath));
        if (!File.Exists(rtfPath)) throw new FileNotFoundException($"RTF file '{rtfPath}' doesn't exist.", rtfPath);

        var effectiveReaderOptions = readerOptions ?? new ReaderOptions();
        var effectiveRtfOptions = ReaderRtfOptionsCloner.CloneOrDefault(rtfOptions);
        ReaderInputLimits.EnforceFileSize(rtfPath, effectiveReaderOptions.MaxInputBytes);
        var source = BuildSourceMetadataFromPath(rtfPath, effectiveReaderOptions.ComputeHashes);

        RtfReadResult readResult = RtfDocument.Load(rtfPath, ReaderRtfOptions.CloneReadOptions(effectiveRtfOptions.RtfReadOptions), encoding);
        IReadOnlyList<ReaderChunk> chunks = ReadRtfResultCore(readResult, source, effectiveReaderOptions, effectiveRtfOptions, cancellationToken).ToArray();
        return CompleteChunkResult(chunks, effectiveRtfOptions.Report);
    }

    /// <summary>
    /// Reads an RTF stream and emits normalized chunks.
    /// </summary>
    public static IEnumerable<ReaderChunk> Read(Stream rtfStream, string? sourceName = null, ReaderOptions? readerOptions = null, ReaderRtfOptions? rtfOptions = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        return ReadResult(rtfStream, sourceName, readerOptions, rtfOptions, encoding, cancellationToken).Value;
    }

    /// <summary>Reads an RTF stream and returns normalized chunks with operation-scoped fidelity diagnostics.</summary>
    public static RtfConversionResult<IReadOnlyList<ReaderChunk>> ReadResult(Stream rtfStream, string? sourceName = null, ReaderOptions? readerOptions = null, ReaderRtfOptions? rtfOptions = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        if (rtfStream == null) throw new ArgumentNullException(nameof(rtfStream));
        if (!rtfStream.CanRead) throw new ArgumentException("RTF stream must be readable.", nameof(rtfStream));

        var effectiveReaderOptions = readerOptions ?? new ReaderOptions();
        var effectiveRtfOptions = ReaderRtfOptionsCloner.CloneOrDefault(rtfOptions);
        var logicalSourceName = NormalizeLogicalSourceName(sourceName, "document.rtf");
        var source = new SourceMetadata {
            Path = logicalSourceName,
            SourceId = BuildSourceId(logicalSourceName)
        };

        Stream parseStream = ReaderInputLimits.EnsureSeekableReadStream(rtfStream, effectiveReaderOptions.MaxInputBytes, cancellationToken, out bool ownsParseStream);
        try {
            long parseStartPosition = parseStream.CanSeek ? parseStream.Position : 0L;
            UpdateSourceMetadataFromSeekableStream(source, parseStream, effectiveReaderOptions.ComputeHashes, parseStartPosition);
            if (parseStream.CanSeek) {
                parseStream.Position = parseStartPosition;
            }

            RtfReadResult readResult = RtfDocument.Load(parseStream, ReaderRtfOptions.CloneReadOptions(effectiveRtfOptions.RtfReadOptions), encoding);
            IReadOnlyList<ReaderChunk> chunks = ReadRtfResultCore(readResult, source, effectiveReaderOptions, effectiveRtfOptions, cancellationToken).ToArray();
            return CompleteChunkResult(chunks, effectiveRtfOptions.Report);
        } finally {
            if (ownsParseStream) {
                parseStream.Dispose();
            }
        }
    }

    /// <summary>
    /// Converts an already loaded RTF semantic model into normalized Reader chunks.
    /// </summary>
    public static IEnumerable<ReaderChunk> Read(RtfDocument document, string sourceName = "document.rtf", ReaderOptions? readerOptions = null, ReaderRtfOptions? rtfOptions = null, CancellationToken cancellationToken = default) {
        return ReadResult(document, sourceName, readerOptions, rtfOptions, cancellationToken).Value;
    }

    /// <summary>Converts an RTF semantic model into normalized chunks with operation-scoped fidelity diagnostics.</summary>
    public static RtfConversionResult<IReadOnlyList<ReaderChunk>> ReadResult(RtfDocument document, string sourceName = "document.rtf", ReaderOptions? readerOptions = null, ReaderRtfOptions? rtfOptions = null, CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (sourceName == null) throw new ArgumentNullException(nameof(sourceName));

        var effectiveReaderOptions = readerOptions ?? new ReaderOptions();
        var effectiveRtfOptions = ReaderRtfOptionsCloner.CloneOrDefault(rtfOptions);
        var logicalSourceName = NormalizeLogicalSourceName(sourceName, "document.rtf");
        var source = new SourceMetadata {
            Path = logicalSourceName,
            SourceId = BuildSourceId(logicalSourceName)
        };

        IReadOnlyList<ReaderChunk> chunks = Read(document, source, effectiveReaderOptions, effectiveRtfOptions, Array.Empty<RtfDiagnostic>(), cancellationToken).ToArray();
        return CompleteChunkResult(chunks, effectiveRtfOptions.Report);
    }

    private static IEnumerable<ReaderChunk> ReadRtfResultCore(RtfReadResult readResult, SourceMetadata source, ReaderOptions readerOptions, ReaderRtfOptions rtfOptions, CancellationToken cancellationToken) {
        if (readResult == null) throw new ArgumentNullException(nameof(readResult));
        rtfOptions.Report.AddReadDiagnostics(readResult.Diagnostics, source.Path);
        IReadOnlyList<RtfDiagnostic> diagnostics = rtfOptions.IncludeDiagnostics ? readResult.Diagnostics : Array.Empty<RtfDiagnostic>();
        foreach (var chunk in Read(readResult.Document, source, readerOptions, rtfOptions, diagnostics, cancellationToken)) {
            yield return chunk;
        }
    }

    private static IEnumerable<ReaderChunk> Read(RtfDocument document, SourceMetadata source, ReaderOptions readerOptions, ReaderRtfOptions rtfOptions, IReadOnlyList<RtfDiagnostic> diagnostics, CancellationToken cancellationToken) {
        int maxChars = readerOptions.MaxChars > 0 ? readerOptions.MaxChars : 8_000;
        var blocks = BuildBlocks(document, source.Path, rtfOptions).ToList();
        if (blocks.Count == 0) {
            blocks.Add(new RtfReaderBlock("warning", 0, "RTF content produced no readable semantic text.", null, null, null, new[] { "RTF content produced no readable semantic text." }));
        }

        if (!rtfOptions.ChunkByBlock) {
            blocks = new List<RtfReaderBlock> {
                CombineBlocks(blocks)
            };
        }

        int emittedIndex = 0;
        IReadOnlyList<string>? documentWarnings = BuildDiagnosticWarnings(diagnostics);
        int documentLinkCount = CountHyperlinkRuns(document);
        int documentFormFieldCount = CountFormFields(document);
        for (int i = 0; i < blocks.Count; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            RtfReaderBlock block = blocks[i];
            var parts = SplitText(block.Markdown ?? block.Text, maxChars);
            if (parts.Count == 0) {
                parts = new[] { string.Empty };
            }

            for (int partIndex = 0; partIndex < parts.Count; partIndex++) {
                cancellationToken.ThrowIfCancellationRequested();
                var warnings = MergeWarnings(partIndex == 0 ? documentWarnings : null, block.Warnings, parts.Count > 1 ? "RTF content was split due to MaxChars." : null);
                string id = BuildChunkId(block.Kind, block.SourceBlockIndex, partIndex, parts.Count);
                string markdown = parts[partIndex];
                yield return EnrichChunk(new ReaderChunk {
                    Id = id,
                    Kind = ReaderInputKind.Rtf,
                    Location = new ReaderLocation {
                        Path = source.Path,
                        BlockIndex = emittedIndex,
                        SourceBlockIndex = block.SourceBlockIndex,
                        SourceBlockKind = block.Kind,
                        BlockAnchor = BuildBlockAnchor(block.Kind, block.SourceBlockIndex, partIndex, parts.Count)
                    },
                    Text = block.Text,
                    Markdown = markdown,
                    Tables = partIndex == 0 ? block.Tables : null,
                    Visuals = partIndex == 0 ? block.Visuals : null,
                    Diagnostics = BuildDiagnostics(block, documentLinkCount, documentFormFieldCount),
                    Warnings = warnings
                }, source, readerOptions.ComputeHashes);

                emittedIndex++;
            }
        }
    }

    private static IEnumerable<RtfReaderBlock> BuildBlocks(RtfDocument document, string sourcePath, ReaderRtfOptions options) {
        int index = 0;
        for (int i = 0; i < document.Blocks.Count; i++) {
            IRtfBlock block = document.Blocks[i];
            switch (block) {
                case RtfParagraph paragraph:
                    yield return BuildParagraphBlock(paragraph, "paragraph", index++);
                    break;
                case RtfTable table:
                    yield return BuildTableBlock(table, index++);
                    break;
                case RtfImage image when options.IncludeImagePlaceholders:
                    yield return BuildImageBlock(image, sourcePath, index++);
                    break;
                case RtfImage:
                    options.Report.Add(RtfConversionSeverity.Warning, "ReaderRtfImageOmitted", "RTF image placeholder was omitted because IncludeImagePlaceholders is false.", RtfConversionAction.Omitted, sourcePath, "image");
                    break;
                case RtfObject rtfObject:
                    options.Report.Add(RtfConversionSeverity.Warning, "ReaderRtfObjectFlattened", "RTF object was flattened to visible text for Reader output.", RtfConversionAction.Flattened, sourcePath, "object");
                    yield return new RtfReaderBlock("object", index++, rtfObject.ToPlainText(), rtfObject.ToPlainText(), null, null, null);
                    break;
                case RtfShape shape:
                    options.Report.Add(RtfConversionSeverity.Warning, "ReaderRtfShapeFlattened", "RTF shape was flattened to visible text for Reader output.", RtfConversionAction.Flattened, sourcePath, "shape");
                    yield return new RtfReaderBlock("shape", index++, shape.ToPlainText(), shape.ToPlainText(), null, null, null);
                    break;
            }
        }

        if (options.IncludeHeadersAndFooters) {
            for (int i = 0; i < document.HeaderFooters.Count; i++) {
                RtfHeaderFooter headerFooter = document.HeaderFooters[i];
                string text = headerFooter.ToPlainText();
                if (string.IsNullOrWhiteSpace(text)) continue;
                yield return new RtfReaderBlock("header-footer", index++, text, text, null, null, null);
            }
        } else if (document.HeaderFooters.Count > 0) {
            options.Report.Add(RtfConversionSeverity.Warning, "ReaderRtfHeaderFooterOmitted", "RTF headers and footers were omitted by Reader options.", RtfConversionAction.Omitted, sourcePath, "header-footer", document.HeaderFooters.Count);
        }

        if (options.IncludeNotes) {
            for (int i = 0; i < document.Notes.Count; i++) {
                RtfNote note = document.Notes[i];
                string text = note.ToPlainText();
                if (string.IsNullOrWhiteSpace(text)) continue;
                yield return new RtfReaderBlock("note", index++, text, text, null, null, null);
            }
        } else if (document.Notes.Count > 0) {
            options.Report.Add(RtfConversionSeverity.Warning, "ReaderRtfNotesOmitted", "RTF notes were omitted by Reader options.", RtfConversionAction.Omitted, sourcePath, "note", document.Notes.Count);
        }
    }

    private static RtfConversionResult<IReadOnlyList<ReaderChunk>> CompleteChunkResult(IReadOnlyList<ReaderChunk> chunks, RtfConversionReport report) {
        string[] adapterWarnings = report.Diagnostics
            .Where(static diagnostic => diagnostic.Code.StartsWith("ReaderRtf", StringComparison.Ordinal))
            .Select(static diagnostic => diagnostic.Code + ": " + diagnostic.Message)
            .ToArray();
        if (adapterWarnings.Length > 0 && chunks.Count > 0) {
            chunks[0].Warnings = MergeWarnings(chunks[0].Warnings, adapterWarnings);
        }

        return new RtfConversionResult<IReadOnlyList<ReaderChunk>>(chunks, report);
    }

    private static RtfReaderBlock BuildParagraphBlock(RtfParagraph paragraph, string kind, int index) {
        string text = paragraph.ToPlainText();
        string markdown = ApplyParagraphMarkdown(paragraph, text);
        return new RtfReaderBlock(kind, index, text, markdown, null, null, null);
    }

    private static string ApplyParagraphMarkdown(RtfParagraph paragraph, string text) {
        if (string.IsNullOrEmpty(text)) return string.Empty;
        if (paragraph.ListKind == RtfListKind.Bullet) return "- " + text;
        if (paragraph.ListKind == RtfListKind.Decimal) return "1. " + text;
        return text;
    }

    private static RtfReaderBlock BuildTableBlock(RtfTable table, int index) {
        ReaderTable readerTable = BuildReaderTable(table, index);
        string markdown = BuildTableMarkdown(readerTable);
        return new RtfReaderBlock("table", index, markdown, markdown, new[] { readerTable }, null, null);
    }

    private static ReaderTable BuildReaderTable(RtfTable table, int index) {
        int columnCount = table.Rows.Count == 0 ? 0 : table.Rows.Max(static row => row.Cells.Count);
        var rows = new List<IReadOnlyList<string>>();
        for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
            RtfTableRow row = table.Rows[rowIndex];
            var values = new List<string>(columnCount);
            for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                string text = columnIndex < row.Cells.Count ? CellText(row.Cells[columnIndex]) : string.Empty;
                values.Add(text);
            }

            rows.Add(values);
        }

        IReadOnlyList<string> columns = columnCount == 0
            ? Array.Empty<string>()
            : Enumerable.Range(1, columnCount).Select(i => "Column " + i.ToString(CultureInfo.InvariantCulture)).ToArray();

        return new ReaderTable {
            Title = "RTF table " + index.ToString(CultureInfo.InvariantCulture),
            Kind = "rtf-table",
            Columns = columns,
            Rows = rows,
            TotalRowCount = rows.Count,
            Location = new ReaderLocation {
                SourceBlockIndex = index,
                SourceBlockKind = "table",
                TableIndex = index
            }
        };
    }

    private static string CellText(RtfTableCell cell) {
        if (cell.Blocks.Count == 0) return string.Empty;
        return string.Join(" ", cell.Blocks.Select(block => block is RtfParagraph paragraph
            ? paragraph.ToPlainText()
            : block is RtfTable table ? NestedTableText(table) : string.Empty).Where(static text => !string.IsNullOrWhiteSpace(text)));
    }

    private static string NestedTableText(RtfTable table) =>
        string.Join(" / ", table.Rows.Select(row => string.Join(" | ", row.Cells.Select(CellText))));

    private static string BuildTableMarkdown(ReaderTable table) {
        if (table.Columns.Count == 0) return string.Empty;

        var builder = new StringBuilder();
        builder.Append("| ");
        builder.Append(string.Join(" | ", table.Columns.Select(EscapeTableCell)));
        builder.AppendLine(" |");
        builder.Append("| ");
        builder.Append(string.Join(" | ", table.Columns.Select(static _ => "---")));
        builder.AppendLine(" |");
        for (int i = 0; i < table.Rows.Count; i++) {
            builder.Append("| ");
            builder.Append(string.Join(" | ", table.Rows[i].Select(EscapeTableCell)));
            builder.AppendLine(" |");
        }

        return builder.ToString().TrimEnd();
    }

    private static string EscapeTableCell(string value) {
        return (value ?? string.Empty)
            .Replace("\\", "\\\\")
            .Replace("|", "\\|")
            .Replace("\r", " ")
            .Replace("\n", " ");
    }

    private static RtfReaderBlock BuildImageBlock(RtfImage image, string sourcePath, int index) {
        string format = image.Format.ToString().ToLowerInvariant();
        string text = "[RTF image: " + format + ", " + image.Data.Length.ToString(CultureInfo.InvariantCulture) + " bytes]";
        var visual = new ReaderVisual {
            Kind = "image",
            Language = "rtf-image",
            Content = text,
            PayloadHash = ComputeSha256Hex(image.Data),
            SourceName = Path.GetFileName(sourcePath),
            MimeType = GetImageMimeType(image.Format),
            Width = image.SourceWidth,
            Height = image.SourceHeight,
            PlacedWidth = image.DesiredWidthTwips,
            PlacedHeight = image.DesiredHeightTwips,
            PlacementCount = 1,
            HasGeometry = image.DesiredWidthTwips.HasValue || image.DesiredHeightTwips.HasValue,
            Location = new ReaderLocation {
                SourceBlockIndex = index,
                SourceBlockKind = "image",
                BlockAnchor = "rtf-image-" + index.ToString("D4", CultureInfo.InvariantCulture)
            }
        };

        return new RtfReaderBlock("image", index, text, text, null, new[] { visual }, null);
    }

    private static string? GetImageMimeType(RtfImageFormat format) {
        switch (format) {
            case RtfImageFormat.Png:
                return "image/png";
            case RtfImageFormat.Jpeg:
                return "image/jpeg";
            case RtfImageFormat.Emf:
                return "image/emf";
            case RtfImageFormat.Wmf:
                return "image/wmf";
            default:
                return null;
        }
    }
}
