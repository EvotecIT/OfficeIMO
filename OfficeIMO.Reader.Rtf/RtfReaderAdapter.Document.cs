using OfficeIMO.Reader;
using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Diagnostics;
using System.Linq;

namespace OfficeIMO.Reader.Rtf;

internal static partial class RtfReaderAdapter {
    /// <summary>Reads an RTF file into the shared rich document envelope.</summary>
    public static OfficeDocumentReadResult ReadDocument(string rtfPath, ReaderOptions? readerOptions = null, ReaderRtfOptions? rtfOptions = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        if (rtfPath == null) throw new ArgumentNullException(nameof(rtfPath));
        if (rtfPath.Length == 0) throw new ArgumentException("RTF path cannot be empty.", nameof(rtfPath));
        if (!File.Exists(rtfPath)) throw new FileNotFoundException($"RTF file '{rtfPath}' doesn't exist.", rtfPath);
        ReaderOptions effectiveReaderOptions = readerOptions ?? new ReaderOptions();
        ReaderRtfOptions effectiveRtfOptions = ReaderRtfOptionsCloner.CloneOrDefault(rtfOptions);
        ReaderInputLimits.EnforceFileSize(rtfPath, effectiveReaderOptions.MaxInputBytes);
        SourceMetadata source = BuildSourceMetadataFromPath(rtfPath, effectiveReaderOptions.ComputeHashes);
        RtfReadResult readResult = RtfDocument.Load(rtfPath, ReaderRtfOptions.CloneReadOptions(effectiveRtfOptions.RtfReadOptions), encoding);
        return BuildRtfDocumentResult(readResult, source, effectiveReaderOptions, effectiveRtfOptions, cancellationToken);
    }

    /// <summary>Reads an RTF stream into the shared rich document envelope.</summary>
    public static OfficeDocumentReadResult ReadDocument(Stream rtfStream, string? sourceName = null, ReaderOptions? readerOptions = null, ReaderRtfOptions? rtfOptions = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        if (rtfStream == null) throw new ArgumentNullException(nameof(rtfStream));
        if (!rtfStream.CanRead) throw new ArgumentException("RTF stream must be readable.", nameof(rtfStream));
        ReaderOptions effectiveReaderOptions = readerOptions ?? new ReaderOptions();
        ReaderRtfOptions effectiveRtfOptions = ReaderRtfOptionsCloner.CloneOrDefault(rtfOptions);
        string logicalSourceName = NormalizeLogicalSourceName(sourceName, "document.rtf");
        var source = new SourceMetadata { Path = logicalSourceName, SourceId = BuildSourceId(logicalSourceName) };
        Stream parseStream = ReaderInputLimits.EnsureSeekableReadStream(rtfStream, effectiveReaderOptions.MaxInputBytes, cancellationToken, out bool ownsParseStream);
        try {
            long start = parseStream.CanSeek ? parseStream.Position : 0L;
            UpdateSourceMetadataFromSeekableStream(source, parseStream, effectiveReaderOptions.ComputeHashes, start);
            if (parseStream.CanSeek) parseStream.Position = start;
            RtfReadResult readResult = RtfDocument.Load(parseStream, ReaderRtfOptions.CloneReadOptions(effectiveRtfOptions.RtfReadOptions), encoding);
            return BuildRtfDocumentResult(readResult, source, effectiveReaderOptions, effectiveRtfOptions, cancellationToken);
        } finally {
            if (ownsParseStream) parseStream.Dispose();
        }
    }

    /// <summary>Converts a loaded RTF semantic document into the shared rich document envelope.</summary>
    public static OfficeDocumentReadResult ReadDocument(RtfDocument document, string sourceName = "document.rtf", ReaderOptions? readerOptions = null, ReaderRtfOptions? rtfOptions = null, CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (sourceName == null) throw new ArgumentNullException(nameof(sourceName));
        ReaderOptions effectiveReaderOptions = readerOptions ?? new ReaderOptions();
        ReaderRtfOptions effectiveRtfOptions = ReaderRtfOptionsCloner.CloneOrDefault(rtfOptions);
        string logicalSourceName = NormalizeLogicalSourceName(sourceName, "document.rtf");
        var source = new SourceMetadata { Path = logicalSourceName, SourceId = BuildSourceId(logicalSourceName) };
        return BuildRtfDocumentResult(document, Array.Empty<RtfDiagnostic>(), source, effectiveReaderOptions, effectiveRtfOptions, cancellationToken);
    }

    /// <summary>Reads an RTF file into the shared rich document JSON envelope.</summary>
    public static string ReadDocumentJson(string rtfPath, ReaderOptions? readerOptions = null, ReaderRtfOptions? rtfOptions = null, Encoding? encoding = null, bool indented = false, CancellationToken cancellationToken = default) {
        return OfficeDocumentReadResultJson.Serialize(ReadDocument(rtfPath, readerOptions, rtfOptions, encoding, cancellationToken), indented);
    }

    /// <summary>Reads an RTF stream into the shared rich document JSON envelope.</summary>
    public static string ReadDocumentJson(Stream rtfStream, string? sourceName = null, ReaderOptions? readerOptions = null, ReaderRtfOptions? rtfOptions = null, Encoding? encoding = null, bool indented = false, CancellationToken cancellationToken = default) {
        return OfficeDocumentReadResultJson.Serialize(ReadDocument(rtfStream, sourceName, readerOptions, rtfOptions, encoding, cancellationToken), indented);
    }

    /// <summary>Converts a loaded RTF semantic document into the shared rich document JSON envelope.</summary>
    public static string ReadDocumentJson(RtfDocument document, string sourceName = "document.rtf", ReaderOptions? readerOptions = null, ReaderRtfOptions? rtfOptions = null, bool indented = false, CancellationToken cancellationToken = default) {
        return OfficeDocumentReadResultJson.Serialize(ReadDocument(document, sourceName, readerOptions, rtfOptions, cancellationToken), indented);
    }

    private static OfficeDocumentReadResult BuildRtfDocumentResult(RtfReadResult readResult, SourceMetadata source, ReaderOptions readerOptions, ReaderRtfOptions rtfOptions, CancellationToken cancellationToken) {
        return BuildRtfDocumentResult(readResult.Document, readResult.Diagnostics, source, readerOptions, rtfOptions, cancellationToken);
    }

    private static OfficeDocumentReadResult BuildRtfDocumentResult(RtfDocument document, IReadOnlyList<RtfDiagnostic> sourceDiagnostics, SourceMetadata source, ReaderOptions readerOptions, ReaderRtfOptions rtfOptions, CancellationToken cancellationToken) {
        IReadOnlyList<RtfDiagnostic> includedDiagnostics = rtfOptions.IncludeDiagnostics ? sourceDiagnostics : Array.Empty<RtfDiagnostic>();
        ReaderChunk[] chunks = Read(document, source, readerOptions, rtfOptions, includedDiagnostics, cancellationToken).ToArray();
        RtfRichProjection projection = ProjectRtfDocument(document, source.Path, readerOptions.MaxTableRows, rtfOptions, cancellationToken);
        var documentSource = new OfficeDocumentSource {
            Path = source.Path,
            SourceId = source.SourceId,
            SourceHash = source.SourceHash,
            LastWriteUtc = source.LastWriteUtc,
            LengthBytes = source.LengthBytes,
            Title = document.Info.Title,
            Author = document.Info.Author,
            Subject = document.Info.Subject,
            Keywords = document.Info.Keywords
        };
        OfficeDocumentReadResult result = DocumentReaderEngine.CreateDocumentResult(
            chunks,
            ReaderInputKind.Rtf,
            documentSource,
            rtfOptions.IncludePageLocations
                ? new[] {
                    "officeimo.reader.rtf.rich-v5",
                    "officeimo.rtf.semantic-model",
                    "officeimo.reader.rtf.pages.explicit"
                }
                : new[] { "officeimo.reader.rtf.rich-v5", "officeimo.rtf.semantic-model" },
            projection.Assets);
        result.Blocks = projection.Blocks;
        result.Tables = projection.Tables;
        result.Links = projection.Links;
        result.Forms = projection.Forms;
        result.Visuals = projection.Visuals;
        result.Metadata = BuildRtfMetadata(document, projection);
        result.Pages = rtfOptions.IncludePageLocations
            ? BuildRtfPages(document, projection, source.Path, cancellationToken)
            : Array.Empty<OfficeDocumentPage>();
        result.Diagnostics = result.Diagnostics
            .Concat(MapRtfDiagnostics(includedDiagnostics, source.Path))
            .Concat(MapReaderRtfConversionDiagnostics(rtfOptions.Report, source.Path))
            .Concat(rtfOptions.IncludePageLocations
                ? BuildRtfPageDiagnostics(document, result.Pages, source.Path)
                : Array.Empty<OfficeDocumentDiagnostic>())
            .ToArray();
        return result;
    }

    private static RtfRichProjection ProjectRtfDocument(RtfDocument document, string sourcePath, int maxTableRows, ReaderRtfOptions options, CancellationToken cancellationToken) {
        var projection = new RtfRichProjection();
        int blockIndex = 0;
        int tableIndex = 0;
        int linkIndex = 0;
        int assetIndex = 0;
        int formIndex = 0;
        foreach (IRtfBlock sourceBlock in document.Blocks) {
            cancellationToken.ThrowIfCancellationRequested();
            string kind = ResolveRtfBlockKind(sourceBlock);
            ReaderLocation location = BuildRtfRichLocation(sourcePath, blockIndex, kind, "rtf-" + kind + "-" + blockIndex.ToString("D4", CultureInfo.InvariantCulture));
            switch (sourceBlock) {
                case RtfParagraph paragraph:
                    projection.Blocks.Add(MapRtfParagraphBlock(paragraph, location));
                    ProjectRtfParagraphContent(paragraph, location, projection, ref linkIndex, ref assetIndex, ref formIndex);
                    break;
                case RtfTable table:
                    ReaderTable mapped = BuildReaderTable(table, tableIndex);
                    mapped.Kind = "rtf-table";
                    mapped.Location = BuildRtfRichLocation(sourcePath, blockIndex, "table", location.BlockAnchor!, tableIndex);
                    ApplyRtfTableLimit(mapped, maxTableRows);
                    projection.Blocks.Add(new OfficeDocumentBlock { Id = location.BlockAnchor!, Kind = "table", Text = BuildRtfTableText(mapped), Location = location });
                    projection.Tables.Add(mapped);
                    foreach (RtfTableRow row in table.Rows) foreach (RtfTableCell cell in row.Cells) foreach (RtfParagraph paragraph in cell.Paragraphs) ProjectRtfParagraphContent(paragraph, location, projection, ref linkIndex, ref assetIndex, ref formIndex);
                    tableIndex++;
                    break;
                case RtfImage image when options.IncludeImagePlaceholders:
                    projection.Blocks.Add(new OfficeDocumentBlock { Id = location.BlockAnchor!, Kind = "image", Text = image.Description ?? "RTF image", Location = location });
                    AddRtfImage(image, location, projection, ref assetIndex);
                    break;
                case RtfObject rtfObject:
                    projection.Blocks.Add(new OfficeDocumentBlock { Id = location.BlockAnchor!, Kind = "object", Text = rtfObject.ToPlainText(), Location = location });
                    ProjectRtfParagraphContent(rtfObject.Result, location, projection, ref linkIndex, ref assetIndex, ref formIndex);
                    if (rtfObject.ResultImage != null) AddRtfImage(rtfObject.ResultImage, location, projection, ref assetIndex);
                    if (rtfObject.Data.Length > 0) projection.Assets.Add(BuildRtfObjectAsset(rtfObject, location, assetIndex++));
                    break;
                case RtfShape shape:
                    projection.Blocks.Add(new OfficeDocumentBlock { Id = location.BlockAnchor!, Kind = "shape", Text = shape.ToPlainText(), Location = location });
                    foreach (RtfParagraph paragraph in shape.TextBoxParagraphs) ProjectRtfParagraphContent(paragraph, location, projection, ref linkIndex, ref assetIndex, ref formIndex);
                    break;
            }
            blockIndex++;
        }
        if (options.IncludeHeadersAndFooters) foreach (RtfHeaderFooter headerFooter in document.HeaderFooters) {
            string id = "rtf-header-footer-" + blockIndex.ToString("D4", CultureInfo.InvariantCulture);
            ReaderLocation location = BuildRtfRichLocation(sourcePath, blockIndex, "header-footer", id);
            projection.Blocks.Add(new OfficeDocumentBlock {
                Id = id,
                Kind = "header-footer",
                Text = headerFooter.ToPlainText(),
                Location = location
            });
            foreach (RtfParagraph paragraph in headerFooter.Paragraphs) {
                ProjectRtfParagraphContent(paragraph, location, projection, ref linkIndex, ref assetIndex, ref formIndex);
            }
            blockIndex++;
        }
        if (options.IncludeNotes) foreach (RtfNote note in document.Notes) {
            string id = "rtf-note-" + blockIndex.ToString("D4", CultureInfo.InvariantCulture);
            ReaderLocation location = BuildRtfRichLocation(sourcePath, blockIndex, "note", id);
            projection.Blocks.Add(new OfficeDocumentBlock {
                Id = id,
                Kind = "note",
                Text = note.ToPlainText(),
                Location = location
            });
            foreach (RtfParagraph paragraph in note.Paragraphs) {
                ProjectRtfParagraphContent(paragraph, location, projection, ref linkIndex, ref assetIndex, ref formIndex);
            }
            blockIndex++;
        }
        return projection;
    }

    private static OfficeDocumentBlock MapRtfParagraphBlock(RtfParagraph paragraph, ReaderLocation location) {
        bool isList = paragraph.ListKind != RtfListKind.None;
        return new OfficeDocumentBlock {
            Id = location.BlockAnchor!, Kind = isList ? "list-item" : paragraph.OutlineLevel.HasValue ? "heading" : "paragraph",
            Text = paragraph.ToPlainText(), Level = paragraph.OutlineLevel ?? paragraph.ListLevel,
            Marker = paragraph.ListKind == RtfListKind.Decimal ? "1." : paragraph.ListKind == RtfListKind.Bullet ? "-" : null,
            Location = location
        };
    }

    private static void ProjectRtfParagraphContent(RtfParagraph paragraph, ReaderLocation location, RtfRichProjection projection, ref int linkIndex, ref int assetIndex, ref int formIndex) {
        foreach (IRtfInline inline in paragraph.Inlines) {
            if (inline is RtfRun run && run.Hyperlink != null) projection.Links.Add(BuildRtfLink(run.Hyperlink.ToString(), null, run.Text, location, linkIndex++));
            else if (inline is RtfField field) {
                RtfHyperlinkFieldInfo? hyperlink = field.HyperlinkField;
                if (hyperlink != null && (hyperlink.Target != null || !string.IsNullOrWhiteSpace(hyperlink.SubAddress))) projection.Links.Add(BuildRtfLink(hyperlink.Target?.ToString(), hyperlink.SubAddress, hyperlink.ScreenTip ?? field.ToPlainText(), location, linkIndex++));
                if (field.FormFieldData != null) projection.Forms.Add(BuildRtfForm(field, location, formIndex++));
                ProjectRtfParagraphContent(field.Result, location, projection, ref linkIndex, ref assetIndex, ref formIndex);
            } else if (inline is RtfImage image) AddRtfImage(image, location, projection, ref assetIndex);
            else if (inline is RtfObject rtfObject) {
                if (rtfObject.ResultImage != null) AddRtfImage(rtfObject.ResultImage, location, projection, ref assetIndex);
                if (rtfObject.Data.Length > 0) projection.Assets.Add(BuildRtfObjectAsset(rtfObject, location, assetIndex++));
            }
        }
    }

    private static void AddRtfImage(RtfImage image, ReaderLocation owner, RtfRichProjection projection, ref int assetIndex) {
        string id = "rtf-image-" + assetIndex.ToString("D4", CultureInfo.InvariantCulture);
        string? mediaType = GetImageMimeType(image.Format);
        string? extension = GetRtfImageExtension(image.Format);
        ReaderLocation location = BuildRtfRichLocation(owner.Path!, owner.SourceBlockIndex, "image", id);
        projection.Assets.Add(new OfficeDocumentAsset {
            Id = id, Kind = "image", MediaType = mediaType, Extension = extension,
            FileName = OfficeDocumentAssetNaming.BuildFileName(id, extension), AltText = image.Description,
            Width = image.SourceWidth, Height = image.SourceHeight, LengthBytes = image.Data.LongLength,
            PayloadHash = ComputeSha256Hex(image.Data), PayloadBytes = image.Data, Location = location
        });
        projection.Visuals.Add(new ReaderVisual {
            Kind = "image", Language = "rtf-image", Content = image.Description ?? id, PayloadHash = ComputeSha256Hex(image.Data),
            MimeType = mediaType, Width = image.SourceWidth, Height = image.SourceHeight,
            PlacedWidth = image.DesiredWidthTwips, PlacedHeight = image.DesiredHeightTwips,
            PlacementCount = 1, HasGeometry = image.DesiredWidthTwips.HasValue || image.DesiredHeightTwips.HasValue, Location = location
        });
        assetIndex++;
    }

    private static OfficeDocumentAsset BuildRtfObjectAsset(RtfObject value, ReaderLocation owner, int index) {
        string id = "rtf-object-" + index.ToString("D4", CultureInfo.InvariantCulture);
        return new OfficeDocumentAsset {
            Id = id, Kind = "embedded-object", FileName = OfficeDocumentAssetNaming.BuildFileName(id, ".bin"),
            Title = value.Name ?? value.ClassName, LengthBytes = value.Data.LongLength,
            PayloadHash = ComputeSha256Hex(value.Data), PayloadBytes = value.Data,
            Location = BuildRtfRichLocation(owner.Path!, owner.SourceBlockIndex, "embedded-object", id)
        };
    }

    private static OfficeDocumentLink BuildRtfLink(string? uri, string? destination, string? text, ReaderLocation owner, int index) => new OfficeDocumentLink {
        Id = "rtf-link-" + index.ToString("D4", CultureInfo.InvariantCulture), Kind = string.IsNullOrWhiteSpace(uri) ? "internal" : "uri",
        Uri = uri, DestinationName = destination, Text = text,
        Location = BuildRtfRichLocation(owner.Path!, owner.SourceBlockIndex, "hyperlink", owner.BlockAnchor + "-link-" + index.ToString("D4", CultureInfo.InvariantCulture))
    };

    private static OfficeDocumentFormField BuildRtfForm(RtfField field, ReaderLocation owner, int index) {
        RtfFormFieldData data = field.FormFieldData!;
        return new OfficeDocumentFormField {
            Id = "rtf-form-" + index.ToString("D4", CultureInfo.InvariantCulture), Name = data.Name,
            Kind = data.Kind?.ToString() ?? "unknown", Value = field.ToPlainText(), IsReadOnly = data.Protected == true,
            Location = BuildRtfRichLocation(owner.Path!, owner.SourceBlockIndex, "form-field", "rtf-form-" + index.ToString("D4", CultureInfo.InvariantCulture))
        };
    }

    private static void ApplyRtfTableLimit(ReaderTable table, int maxRows) {
        table.ColumnProfiles = ReaderTableProfiler.CreateProfiles(table.Columns, table.Rows);
        if (maxRows <= 0 || table.Rows.Count <= maxRows) return;
        table.Rows = table.Rows.Take(maxRows).ToArray();
        table.Truncated = true;
        table.ColumnProfiles = ReaderTableProfiler.CreateProfiles(table.Columns, table.Rows);
    }

    private static string BuildRtfTableText(ReaderTable table) {
        return string.Join(Environment.NewLine, table.Rows.Select(static row => string.Join(" | ", row)));
    }

    private static IEnumerable<OfficeDocumentDiagnostic> MapRtfDiagnostics(IReadOnlyList<RtfDiagnostic> diagnostics, string path) {
        foreach (RtfDiagnostic diagnostic in diagnostics) yield return new OfficeDocumentDiagnostic {
            Severity = diagnostic.Severity == RtfDiagnosticSeverity.Error ? OfficeDocumentDiagnosticSeverity.Error : diagnostic.Severity == RtfDiagnosticSeverity.Info ? OfficeDocumentDiagnosticSeverity.Information : OfficeDocumentDiagnosticSeverity.Warning,
            Category = OfficeDocumentDiagnosticCategory.Parsing, Code = diagnostic.Code, Message = diagnostic.Message,
            Source = "officeimo.rtf", IsRecoverable = diagnostic.Severity != RtfDiagnosticSeverity.Error,
            Location = new ReaderLocation { Path = path, SourceBlockKind = "diagnostic", BlockAnchor = "rtf-position-" + diagnostic.Position.ToString(CultureInfo.InvariantCulture) }
        };
    }

    private static IEnumerable<OfficeDocumentDiagnostic> MapReaderRtfConversionDiagnostics(RtfConversionReport report, string path) {
        foreach (RtfConversionDiagnostic diagnostic in report.Diagnostics) {
            if (!diagnostic.Code.StartsWith("ReaderRtf", StringComparison.Ordinal)) continue;
            yield return new OfficeDocumentDiagnostic {
                Severity = diagnostic.Severity == RtfConversionSeverity.Error
                    ? OfficeDocumentDiagnosticSeverity.Error
                    : diagnostic.Severity == RtfConversionSeverity.Information
                        ? OfficeDocumentDiagnosticSeverity.Information
                        : OfficeDocumentDiagnosticSeverity.Warning,
                Category = OfficeDocumentDiagnosticCategory.Content,
                Code = diagnostic.Code,
                Message = diagnostic.Message,
                Source = "officeimo.reader.rtf",
                IsRecoverable = diagnostic.Severity != RtfConversionSeverity.Error,
                Location = new ReaderLocation {
                    Path = diagnostic.SourcePath ?? path,
                    SourceBlockKind = diagnostic.Feature ?? "adapter",
                    BlockAnchor = "rtf-adapter-" + diagnostic.Code
                },
                Attributes = new Dictionary<string, string> {
                    ["action"] = diagnostic.Action.ToString(),
                    ["count"] = diagnostic.Count.ToString(CultureInfo.InvariantCulture),
                    ["detail"] = diagnostic.Detail ?? string.Empty
                }
            };
        }
    }

    private static IReadOnlyList<OfficeDocumentMetadataEntry> BuildRtfMetadata(RtfDocument document, RtfRichProjection projection) {
        var metadata = new List<OfficeDocumentMetadataEntry> {
            RtfMetadata("rtf-font-count", "FontCount", document.Fonts.Count),
            RtfMetadata("rtf-style-count", "StyleCount", document.Styles.Count),
            RtfMetadata("rtf-block-count", "BlockCount", projection.Blocks.Count),
            RtfMetadata("rtf-table-count", "TableCount", projection.Tables.Count),
            RtfMetadata("rtf-link-count", "LinkCount", projection.Links.Count),
            RtfMetadata("rtf-asset-count", "AssetCount", projection.Assets.Count),
            RtfMetadata("rtf-form-count", "FormFieldCount", projection.Forms.Count)
        };
        if (document.Info.NumberOfPages.HasValue && document.Info.NumberOfPages.Value > 0) {
            metadata.Add(RtfMetadata("rtf-page-count", "NumberOfPages", document.Info.NumberOfPages.Value));
        }
        return metadata.AsReadOnly();
    }

    private static OfficeDocumentMetadataEntry RtfMetadata(string id, string name, int count) => new OfficeDocumentMetadataEntry {
        Id = id, Category = "rtf.summary", Name = name, Value = count.ToString(CultureInfo.InvariantCulture), ValueType = "count"
    };

    private static string ResolveRtfBlockKind(IRtfBlock block) => block is RtfParagraph paragraph && paragraph.ListKind != RtfListKind.None ? "list-item" : block switch {
        RtfParagraph => "paragraph", RtfTable => "table", RtfImage => "image", RtfObject => "object", RtfShape => "shape", _ => "block"
    };

    private static string? GetRtfImageExtension(RtfImageFormat format) => format switch {
        RtfImageFormat.Png => ".png", RtfImageFormat.Jpeg => ".jpg", RtfImageFormat.Emf => ".emf", RtfImageFormat.Wmf => ".wmf", RtfImageFormat.Dib => ".dib", _ => null
    };

    private static ReaderLocation BuildRtfRichLocation(string path, int? blockIndex, string kind, string anchor, int? tableIndex = null) => new ReaderLocation {
        Path = path, SourceBlockIndex = blockIndex, SourceBlockKind = kind, BlockAnchor = anchor, TableIndex = tableIndex
    };

    private sealed class RtfRichProjection {
        internal List<OfficeDocumentBlock> Blocks { get; } = new List<OfficeDocumentBlock>();
        internal List<ReaderTable> Tables { get; } = new List<ReaderTable>();
        internal List<OfficeDocumentAsset> Assets { get; } = new List<OfficeDocumentAsset>();
        internal List<OfficeDocumentLink> Links { get; } = new List<OfficeDocumentLink>();
        internal List<OfficeDocumentFormField> Forms { get; } = new List<OfficeDocumentFormField>();
        internal List<ReaderVisual> Visuals { get; } = new List<ReaderVisual>();
    }
}
