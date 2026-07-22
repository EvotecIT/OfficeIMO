using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using System.Runtime.CompilerServices;

namespace OfficeIMO.Reader;

/// <summary>Loss-aware PDF projection over the shared normalized Reader result.</summary>
public static class OfficeDocumentReadResultPdfExtensions {
    private const string ConverterName = "OfficeIMO.Reader.Pdf";

    /// <summary>
    /// Projects normalized Reader content into a PDF document while merging source and PDF diagnostics.
    /// </summary>
    public static PdfDocumentConversionResult ToPdfDocumentResult(
        this OfficeDocumentReadResult source,
        ReaderPdfProjectionOptions? options = null) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        options ??= new ReaderPdfProjectionOptions();
        options.Validate();

        var report = new PdfConversionReport();
        AddSourceDiagnostics(source, report);
        PdfOptions pdfOptions = options.PdfOptions?.Clone() ?? new PdfOptions();
        pdfOptions.ReportDiagnosticsTo(report, ConverterName);
        PdfDocument document = PdfDocument.Create(pdfOptions)
            .Meta(source.Source.Title, source.Source.Author, source.Source.Subject, source.Source.Keywords);
        OfficeRasterDecodeOptions rasterDecodeOptions = options.SnapshotRasterDecodeOptions();
        var identities = new ProjectionIdentitySet();
        AssetProjectionSummary assetSummary = default;

        if (!string.IsNullOrWhiteSpace(source.Source.Title)) document.H1(source.Source.Title!);
        if (options.IncludeMetadata) ComposeMetadata(document, source.Metadata);

        if (source.Pages.Count > 0) {
            for (int index = 0; index < source.Pages.Count; index++) {
                OfficeDocumentPage page = source.Pages[index];
                if (!string.IsNullOrWhiteSpace(page.Name)) document.H2(page.Name!);
                assetSummary = assetSummary.Combine(ComposeContent(
                    document,
                    identities.TakeBlocks(page.Blocks),
                    identities.TakeTables(page.Tables, page),
                    identities.TakeAssets(page.Assets),
                    identities.TakeLinks(page.Links),
                    identities.TakeForms(page.Forms),
                    options,
                    rasterDecodeOptions,
                    !string.IsNullOrWhiteSpace(pdfOptions.CatalogUriBase),
                    report,
                    BuildSourceLabel(source, page, index)));
                if (options.PagePolicy == ReaderPdfPagePolicy.PreserveSourcePages && index + 1 < source.Pages.Count) document.PageBreak();
            }

            assetSummary = assetSummary.Combine(ComposeContent(
                document,
                identities.TakeBlocks(source.Blocks),
                identities.TakeTables(source.Tables),
                identities.TakeAssets(source.Assets),
                identities.TakeLinks(source.Links),
                identities.TakeForms(source.Forms),
                options,
                rasterDecodeOptions,
                !string.IsNullOrWhiteSpace(pdfOptions.CatalogUriBase),
                report,
                source.Kind + "/document"));
        } else {
            assetSummary = ComposeContent(
                document,
                identities.TakeBlocks(source.Blocks),
                identities.TakeTables(source.Tables),
                identities.TakeAssets(source.Assets),
                identities.TakeLinks(source.Links),
                identities.TakeForms(source.Forms),
                options,
                rasterDecodeOptions,
                !string.IsNullOrWhiteSpace(pdfOptions.CatalogUriBase),
                report,
                source.Kind.ToString());
        }

        AddSourceSpecificPolicyEvidence(source, options, assetSummary, report);
        return new PdfDocumentConversionResult(document, report);
    }

    private static AssetProjectionSummary ComposeContent(
        PdfDocument document,
        IReadOnlyList<OfficeDocumentBlock> blocks,
        IReadOnlyList<ReaderTable> tables,
        IReadOnlyList<OfficeDocumentAsset> assets,
        IReadOnlyList<OfficeDocumentLink> links,
        IReadOnlyList<OfficeDocumentFormField> forms,
        ReaderPdfProjectionOptions options,
        OfficeRasterDecodeOptions rasterDecodeOptions,
        bool allowRelativeUriLinks,
        PdfConversionReport report,
        string sourceLabel) {
        ComposeBlocksAndTables(document, blocks, tables, report, sourceLabel);
        AssetProjectionSummary assetSummary = ComposeAssets(document, assets, options.AssetPolicy, rasterDecodeOptions, report, sourceLabel);
        ComposeLinks(document, links, options.LinkPolicy, allowRelativeUriLinks, report, sourceLabel);
        ComposeForms(document, forms, options.FormPolicy, report, sourceLabel);
        return assetSummary;
    }

    private static void ComposeBlocksAndTables(
        PdfDocument document,
        IReadOnlyList<OfficeDocumentBlock> blocks,
        IReadOnlyList<ReaderTable> tables,
        PdfConversionReport report,
        string sourceLabel) {
        var matchedBlocks = new HashSet<OfficeDocumentBlock>(ReferenceIdentityComparer<OfficeDocumentBlock>.Instance);
        var items = new List<ProjectionContentItem>(blocks.Count + tables.Count);
        for (int tableIndex = 0; tableIndex < tables.Count; tableIndex++) {
            ReaderTable table = tables[tableIndex];
            OfficeDocumentBlock? correlated = FindCorrelatedTableBlock(blocks, table, out int correlatedIndex);
            if (correlated != null) matchedBlocks.Add(correlated);
            items.Add(ProjectionContentItem.ForTable(
                table,
                tableIndex,
                correlated?.Location,
                correlatedIndex >= 0 ? correlatedIndex : blocks.Count + tableIndex));
        }
        for (int blockIndex = 0; blockIndex < blocks.Count; blockIndex++) {
            OfficeDocumentBlock block = blocks[blockIndex];
            if (matchedBlocks.Contains(block)) continue;
            items.Add(ProjectionContentItem.ForBlock(block, blockIndex));
        }

        foreach (ProjectionContentItem item in items.OrderBy(static item => item.Position).ThenBy(static item => item.InsertionIndex)) {
            if (item.Block != null) ComposeBlock(document, item.Block);
            else ComposeTable(document, item.Table!, item.TableIndex, report, sourceLabel);
        }
    }

    private static void ComposeBlock(PdfDocument document, OfficeDocumentBlock block) {
        if (string.IsNullOrWhiteSpace(block.Text)) return;
        string kind = block.Kind ?? string.Empty;
        if (kind.IndexOf("heading", StringComparison.OrdinalIgnoreCase) >= 0) {
            int level = block.Level ?? 2;
            if (level <= 1) document.H1(block.Text);
            else if (level == 2) document.H2(block.Text);
            else document.H3(block.Text);
        } else if (kind.IndexOf("list", StringComparison.OrdinalIgnoreCase) >= 0) {
            document.Bullets(new[] { block.Text });
        } else {
            document.Paragraph(paragraph => paragraph.Text(block.Text));
        }
    }

    private static void ComposeTable(PdfDocument document, ReaderTable table, int tableIndex, PdfConversionReport report, string sourceLabel) {
        if (!string.IsNullOrWhiteSpace(table.Title)) document.H3(table.Title!);
        var rows = new List<string[]>();
        if (table.Columns.Count > 0) rows.Add(table.Columns.ToArray());
        rows.AddRange(table.Rows.Select(static row => row.ToArray()));
        if (rows.Count > 0) document.Table(rows, style: new PdfTableStyle { HeaderRowCount = table.Columns.Count > 0 ? 1 : 0 });
        if (table.Truncated) {
            report.Add(Warning("reader-table-truncated", sourceLabel + "/table-" + tableIndex, "The normalized source table was truncated before PDF projection."));
        }
    }

    private static AssetProjectionSummary ComposeAssets(
        PdfDocument document,
        IReadOnlyList<OfficeDocumentAsset> assets,
        ReaderPdfAssetPolicy policy,
        OfficeRasterDecodeOptions rasterDecodeOptions,
        PdfConversionReport report,
        string sourceLabel) {
        if (assets.Count == 0) return default;
        int imageCandidates = assets.Count(IsImageCandidate);
        if (policy == ReaderPdfAssetPolicy.Omit) {
            report.Add(Warning("reader-assets-omitted", sourceLabel, assets.Count + " normalized assets were omitted by policy."));
            return new AssetProjectionSummary(imageCandidates, embedded: 0, listed: 0, omitted: assets.Count);
        }

        document.H3("Assets");
        int embedded = 0;
        int listed = 0;
        foreach (OfficeDocumentAsset asset in assets) {
            string label = asset.Title ?? asset.FileName ?? asset.Id ?? "asset";
            if (policy == ReaderPdfAssetPolicy.EmbedSupportedImages &&
                asset.PayloadBytes != null &&
                TryPreparePdfImage(asset.PayloadBytes, rasterDecodeOptions, report, sourceLabel + "/" + label, out byte[] imageBytes, out OfficeImageInfo? info)) {
                double width = Math.Min(480D, Math.Max(36D, (info?.Width ?? asset.Width ?? 320) * 0.75D));
                double height = Math.Max(24D, (info?.Height ?? asset.Height ?? 180) * width / Math.Max(1D, info?.Width ?? asset.Width ?? 320));
                document.Image(imageBytes, width, height, label);
                embedded++;
                continue;
            }

            document.Paragraph(paragraph => paragraph.Bold(label).Text(BuildAssetSuffix(asset)));
            listed++;
            if (policy == ReaderPdfAssetPolicy.EmbedSupportedImages && asset.PayloadBytes != null) {
                report.Add(Warning("reader-asset-listed-not-embedded", sourceLabel + "/" + label, "The asset payload is not part of the shared PDF raster contract and was represented as metadata."));
            }
        }
        return new AssetProjectionSummary(imageCandidates, embedded, listed, omitted: 0);
    }

    private static bool TryPreparePdfImage(
        byte[] sourceBytes,
        OfficeRasterDecodeOptions rasterDecodeOptions,
        PdfConversionReport report,
        string sourceLabel,
        out byte[] imageBytes,
        out OfficeImageInfo? imageInfo) {
        imageBytes = Array.Empty<byte>();
        imageInfo = null;
        bool identified = OfficeImageReader.TryIdentify(sourceBytes, null, out OfficeImageInfo identifiedInfo);
        bool isDirectPdfImage = identifiedInfo.Format == OfficeImageFormat.Jpeg ||
            (identifiedInfo.Format == OfficeImageFormat.Png &&
             OfficePngReader.TryGetFrameCount(sourceBytes, out int pngFrameCount) &&
             pngFrameCount == 1);
        if (identified &&
            rasterDecodeOptions.FrameIndex == 0 &&
            isDirectPdfImage &&
            OfficeImagePdfCompatibility.TryValidate(sourceBytes, out imageInfo, out _)) {
            imageBytes = sourceBytes;
            return true;
        }

        if (!identified ||
            !OfficeImagePdfCompatibility.TryValidateTranscodeDimensions(
                identifiedInfo,
                OfficeImagePdfCompatibility.DefaultMaximumTranscodePixels,
                out _)) return false;

        bool converted = OfficeImagePngConverter.TryConvertToPng(
            sourceBytes,
            rasterDecodeOptions,
            out byte[] normalizedPng,
            out OfficeRasterDecodeInfo decodeInfo);
        if (decodeInfo.IsAnimated) {
            if (converted && decodeInfo.AnimationDiscarded) {
                report.Add(Warning(
                    "reader-asset-animation-frame-selected",
                    sourceLabel,
                    "Frame " + decodeInfo.SelectedFrameIndex + " of " + decodeInfo.FrameCount + " was embedded as a static image; remaining animation frames were not retained."));
            } else {
                string code = rasterDecodeOptions.AnimationPolicy == OfficeRasterAnimationPolicy.RejectAnimated
                    ? "reader-asset-animation-rejected"
                    : "reader-asset-animation-not-supported";
                report.Add(Warning(code, sourceLabel, decodeInfo.Diagnostic ?? "The animated raster asset could not be represented as a static PDF image."));
            }
        }
        if (!converted || !OfficeImagePdfCompatibility.TryValidate(normalizedPng, out OfficeImageInfo? normalizedInfo, out _)) return false;

        imageBytes = normalizedPng;
        imageInfo = identified ? identifiedInfo : normalizedInfo;
        return true;
    }

    private static OfficeDocumentBlock? FindCorrelatedTableBlock(
        IReadOnlyList<OfficeDocumentBlock> blocks,
        ReaderTable table,
        out int blockIndex) {
        for (int index = 0; index < blocks.Count; index++) {
            OfficeDocumentBlock block = blocks[index];
            if (IsTableBlock(block) && TableMatchesBlock(table, block)) {
                blockIndex = index;
                return block;
            }
        }
        blockIndex = -1;
        return null;
    }

    private static bool IsTableBlock(OfficeDocumentBlock block) =>
        string.Equals(block.Kind, "table", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(block.Location?.SourceBlockKind, "table", StringComparison.OrdinalIgnoreCase);

    private static bool TableMatchesBlock(ReaderTable table, OfficeDocumentBlock block) {
        ReaderLocation? tableLocation = table.Location;
        ReaderLocation? blockLocation = block.Location;
        string? tableAnchor = tableLocation?.BlockAnchor;
        if (!string.IsNullOrWhiteSpace(tableAnchor) &&
            (string.Equals(tableAnchor, block.Id, StringComparison.Ordinal) ||
             string.Equals(tableAnchor, blockLocation?.BlockAnchor, StringComparison.Ordinal))) return true;
        if (tableLocation?.TableIndex.HasValue == true &&
            blockLocation?.TableIndex.HasValue == true &&
            tableLocation.TableIndex == blockLocation.TableIndex &&
            SameContainer(tableLocation, blockLocation)) return true;
        return tableLocation?.SourceBlockIndex.HasValue == true &&
               blockLocation?.SourceBlockIndex.HasValue == true &&
               tableLocation.SourceBlockIndex == blockLocation.SourceBlockIndex &&
               SameContainer(tableLocation, blockLocation);
    }

    private static bool SameContainer(ReaderLocation left, ReaderLocation right) =>
        left.Page == right.Page &&
        left.Slide == right.Slide &&
        string.Equals(left.Sheet, right.Sheet, StringComparison.Ordinal) &&
        string.Equals(left.Path, right.Path, StringComparison.Ordinal);

    private static void ComposeLinks(
        PdfDocument document,
        IReadOnlyList<OfficeDocumentLink> links,
        ReaderPdfLinkPolicy policy,
        bool allowRelativeUriLinks,
        PdfConversionReport report,
        string sourceLabel) {
        if (links.Count == 0) return;
        if (policy == ReaderPdfLinkPolicy.Omit) {
            report.Add(Warning("reader-links-omitted", sourceLabel, links.Count + " normalized links were omitted by policy."));
            return;
        }

        document.H3("Links");
        foreach (OfficeDocumentLink link in links) {
            string text = link.Text ?? link.Uri ?? link.DestinationName ?? link.RemoteFile ?? link.Id ?? "link";
            bool canPreserveUri = Uri.TryCreate(link.Uri, UriKind.RelativeOrAbsolute, out Uri? uri) &&
                (uri.IsAbsoluteUri || allowRelativeUriLinks);
            if (policy == ReaderPdfLinkPolicy.PreserveUriLinks && canPreserveUri) {
                document.Paragraph(paragraph => paragraph.Link(text, link.Uri!));
            } else {
                string target = link.DestinationName ?? link.RemoteDestinationName ?? link.RemoteFile ?? link.Uri ?? string.Empty;
                document.Paragraph(paragraph => paragraph.Text(string.IsNullOrWhiteSpace(target) ? text : text + ": " + target));
                if (policy == ReaderPdfLinkPolicy.PreserveUriLinks) {
                    report.Add(Warning("reader-navigation-listed", sourceLabel + "/" + text, "A non-URI navigation target was represented as text."));
                }
            }
        }
    }

    private static void ComposeForms(PdfDocument document, IReadOnlyList<OfficeDocumentFormField> forms, ReaderPdfFormPolicy policy, PdfConversionReport report, string sourceLabel) {
        if (forms.Count == 0) return;
        if (policy == ReaderPdfFormPolicy.Omit) {
            report.Add(Warning("reader-forms-omitted", sourceLabel, forms.Count + " normalized form fields were omitted by policy."));
            return;
        }
        document.H3("Form values");
        document.Table(forms.Select(static field => new[] { field.Name ?? field.Id, field.Value ?? string.Empty, field.Kind }), style: new PdfTableStyle { HeaderRowCount = 0 });
        report.Add(Warning("reader-forms-rendered-static", sourceLabel, "Source form fields were rendered as current values rather than recreated as interactive controls."));
    }

    private static void ComposeMetadata(PdfDocument document, IReadOnlyList<OfficeDocumentMetadataEntry> metadata) {
        if (metadata.Count == 0) return;
        string[][] rows = metadata
            .Where(static entry => !string.IsNullOrWhiteSpace(entry.Name))
            .Select(static entry => new[] { entry.Name, entry.Value ?? string.Empty })
            .ToArray();
        if (rows.Length > 0) {
            document.Table(rows, style: new PdfTableStyle { HeaderRowCount = 0, RowStripeFill = null, SpacingAfter = 10 });
        }
    }

    private static void AddSourceDiagnostics(OfficeDocumentReadResult source, PdfConversionReport report) {
        foreach (OfficeDocumentDiagnostic diagnostic in source.Diagnostics) {
            report.Add(new PdfConversionWarning(
                ConverterName,
                string.IsNullOrWhiteSpace(diagnostic.Code) ? "reader-source-diagnostic" : diagnostic.Code,
                diagnostic.Source ?? source.Kind.ToString(),
                diagnostic.Message,
                diagnostic.Severity == OfficeDocumentDiagnosticSeverity.Error
                    ? PdfConversionWarningSeverity.Error
                    : diagnostic.Severity == OfficeDocumentDiagnosticSeverity.Information
                        ? PdfConversionWarningSeverity.Information
                        : PdfConversionWarningSeverity.Warning,
                details: diagnostic.Attributes));
        }
    }

    private static void AddSourceSpecificPolicyEvidence(
        OfficeDocumentReadResult source,
        ReaderPdfProjectionOptions options,
        AssetProjectionSummary assetSummary,
        PdfConversionReport report) {
        if (source.Kind == ReaderInputKind.Email) {
            report.Add(Information("reader-email-policy", "Email", "Email body content follows normalized block order; attachments follow the configured asset policy (" + options.AssetPolicy + ")."));
        } else if (source.Kind == ReaderInputKind.Epub) {
            report.Add(Information("reader-epub-policy", "EPUB", "Book resources follow the configured asset policy and normalized chapters follow the configured page policy (" + options.PagePolicy + ")."));
        } else if (source.Kind == ReaderInputKind.Visio) {
            if (assetSummary.ImageCandidates == 0) {
                report.Add(Warning("reader-visio-semantic-fallback", "Visio", "No raster preview was available, so diagram pages were projected through normalized semantic content."));
            } else if (options.AssetPolicy == ReaderPdfAssetPolicy.EmbedSupportedImages && assetSummary.Embedded > 0) {
                report.Add(Information("reader-visio-preview-embedded", "Visio", assetSummary.Embedded + " diagram preview image(s) were embedded; semantic page content remains searchable."));
            } else if (options.AssetPolicy == ReaderPdfAssetPolicy.ListMetadata) {
                report.Add(Information("reader-visio-preview-listed", "Visio", "Diagram preview assets were listed as metadata by policy and were not embedded; semantic page content remains searchable."));
            } else if (options.AssetPolicy == ReaderPdfAssetPolicy.Omit) {
                report.Add(Information("reader-visio-preview-omitted", "Visio", "Diagram preview assets were omitted by policy; semantic page content remains searchable."));
            } else {
                report.Add(Warning("reader-visio-preview-not-embedded", "Visio", "Diagram preview assets were available but could not be embedded through the configured shared raster policy; semantic page content remains searchable."));
            }
        }
    }

    private static bool IsImageCandidate(OfficeDocumentAsset asset) =>
        asset.Kind.IndexOf("image", StringComparison.OrdinalIgnoreCase) >= 0 ||
        asset.Kind.IndexOf("preview", StringComparison.OrdinalIgnoreCase) >= 0 ||
        (asset.MediaType?.StartsWith("image/", StringComparison.OrdinalIgnoreCase) ?? false) ||
        (asset.PayloadBytes != null && OfficeImageReader.TryIdentify(asset.PayloadBytes, null, out _));

    private static string BuildSourceLabel(OfficeDocumentReadResult source, OfficeDocumentPage page, int index) =>
        source.Kind + "/" + (page.Name ?? "page-" + (page.Number ?? index + 1).ToString(System.Globalization.CultureInfo.InvariantCulture));

    private static string BuildAssetSuffix(OfficeDocumentAsset asset) {
        var parts = new List<string>();
        if (!string.IsNullOrWhiteSpace(asset.MediaType)) parts.Add(asset.MediaType!);
        if (asset.LengthBytes.HasValue) parts.Add(asset.LengthBytes.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) + " bytes");
        return parts.Count == 0 ? string.Empty : " (" + string.Join(", ", parts) + ")";
    }

    private static PdfConversionWarning Warning(string code, string source, string message) =>
        new PdfConversionWarning(ConverterName, code, source, message);

    private static PdfConversionWarning Information(string code, string source, string message) =>
        new PdfConversionWarning(ConverterName, code, source, message, PdfConversionWarningSeverity.Information);

    private readonly struct AssetProjectionSummary {
        internal AssetProjectionSummary(int imageCandidates, int embedded, int listed, int omitted) {
            ImageCandidates = imageCandidates;
            Embedded = embedded;
            Listed = listed;
            Omitted = omitted;
        }

        internal int ImageCandidates { get; }
        internal int Embedded { get; }
        internal int Listed { get; }
        internal int Omitted { get; }

        internal AssetProjectionSummary Combine(AssetProjectionSummary other) =>
            new AssetProjectionSummary(
                ImageCandidates + other.ImageCandidates,
                Embedded + other.Embedded,
                Listed + other.Listed,
                Omitted + other.Omitted);
    }

    private readonly struct ProjectionContentItem {
        private ProjectionContentItem(
            OfficeDocumentBlock? block,
            ReaderTable? table,
            int tableIndex,
            ReaderLocation? location,
            int insertionIndex) {
            Block = block;
            Table = table;
            TableIndex = tableIndex;
            Position = GetPosition(location);
            InsertionIndex = insertionIndex;
        }

        internal OfficeDocumentBlock? Block { get; }
        internal ReaderTable? Table { get; }
        internal int TableIndex { get; }
        internal int Position { get; }
        internal int InsertionIndex { get; }

        internal static ProjectionContentItem ForBlock(OfficeDocumentBlock block, int insertionIndex) =>
            new ProjectionContentItem(block, null, -1, block.Location, insertionIndex);

        internal static ProjectionContentItem ForTable(
            ReaderTable table,
            int tableIndex,
            ReaderLocation? correlatedLocation,
            int insertionIndex) =>
            new ProjectionContentItem(null, table, tableIndex, correlatedLocation ?? table.Location, insertionIndex);

        private static int GetPosition(ReaderLocation? location) =>
            location?.SourceBlockIndex
            ?? location?.BlockIndex
            ?? location?.StartLine
            ?? location?.NormalizedStartLine
            ?? int.MaxValue;
    }

    private sealed class ProjectionIdentitySet {
        private readonly HashSet<string> _blocks = new HashSet<string>(StringComparer.Ordinal);
        private readonly HashSet<string> _tables = new HashSet<string>(StringComparer.Ordinal);
        private readonly HashSet<string> _assets = new HashSet<string>(StringComparer.Ordinal);
        private readonly HashSet<string> _links = new HashSet<string>(StringComparer.Ordinal);
        private readonly HashSet<string> _forms = new HashSet<string>(StringComparer.Ordinal);

        internal IReadOnlyList<OfficeDocumentBlock> TakeBlocks(IReadOnlyList<OfficeDocumentBlock> items) =>
            Take(items, _blocks, OfficeDocumentModelTraversal.BuildBlockIdentity);

        internal IReadOnlyList<ReaderTable> TakeTables(IReadOnlyList<ReaderTable> items) =>
            Take(items, _tables, static item => OfficeDocumentModelTraversal.BuildTableIdentity(item));

        internal IReadOnlyList<ReaderTable> TakeTables(IReadOnlyList<ReaderTable> items, OfficeDocumentPage page) {
            var result = new List<ReaderTable>(items.Count);
            for (int index = 0; index < items.Count; index++) {
                ReaderTable item = items[index];
                if (item != null && _tables.Add(OfficeDocumentModelTraversal.BuildTableIdentity(item, page, index))) result.Add(item);
            }
            return result;
        }

        internal IReadOnlyList<OfficeDocumentAsset> TakeAssets(IReadOnlyList<OfficeDocumentAsset> items) =>
            Take(items, _assets, OfficeDocumentModelTraversal.BuildAssetIdentity);

        internal IReadOnlyList<OfficeDocumentLink> TakeLinks(IReadOnlyList<OfficeDocumentLink> items) =>
            Take(items, _links, OfficeDocumentModelTraversal.BuildLinkIdentity);

        internal IReadOnlyList<OfficeDocumentFormField> TakeForms(IReadOnlyList<OfficeDocumentFormField> items) =>
            Take(items, _forms, OfficeDocumentModelTraversal.BuildFormIdentity);

        private static IReadOnlyList<T> Take<T>(IReadOnlyList<T> items, HashSet<string> seen, Func<T, string> identity) where T : class {
            var result = new List<T>(items.Count);
            foreach (T item in items) {
                if (item != null && seen.Add(identity(item))) result.Add(item);
            }
            return result;
        }
    }

    private sealed class ReferenceIdentityComparer<T> : IEqualityComparer<T> where T : class {
        internal static ReferenceIdentityComparer<T> Instance { get; } = new ReferenceIdentityComparer<T>();

        private ReferenceIdentityComparer() {
        }

        public bool Equals(T? x, T? y) => ReferenceEquals(x, y);
        public int GetHashCode(T obj) => RuntimeHelpers.GetHashCode(obj);
    }
}
