using System.Globalization;
using System.Text;
using Microsoft.Extensions.AI;
using Microsoft.Extensions.Documents;

namespace OfficeIMO.Reader.DocumentExtraction;

/// <summary>Projects one OfficeIMO read result onto the shared extraction document model.</summary>
internal sealed class ExtractionDocumentMapper {
    private const string ColumnHeaderRole = "columnHeader";

    private readonly OfficeDocumentReadResult _result;
    private readonly OfficeIMODocumentExtractionOptions _options;
    private readonly ExtractionDocument _document;
    private readonly bool _hasPdfGeometry;
    private readonly Dictionary<string, ExtractionDocumentSourcePage> _frames = new(StringComparer.Ordinal);
    private readonly Dictionary<string, OfficeDocumentPage> _pagesByKey = new(StringComparer.Ordinal);
    private readonly List<(int Level, ExtractionDocumentSection Section)> _headings = new();
    private readonly Dictionary<string, OfficeDocumentBlock> _ocrTextByAssetId = new(StringComparer.Ordinal);
    private readonly HashSet<string> _consumedOcrBlockIds = new(StringComparer.Ordinal);
    private readonly HashSet<string> _tableAnchors = new(StringComparer.Ordinal);
    private readonly HashSet<ReaderTable> _duplicateTables = new();
    private readonly Dictionary<OfficeDocumentBlock, OfficeDocumentAsset> _imagesForBlocks = new();
    private ExtractionDocumentSection? _preamble;

    private ExtractionDocumentMapper(OfficeDocumentReadResult result, string identifier, OfficeIMODocumentExtractionOptions options) {
        _result = result;
        _options = options;
        _hasPdfGeometry = result.Kind == ReaderInputKind.Pdf;
        _document = new ExtractionDocument(identifier, null) { RawRepresentation = result };
    }

    internal static ExtractionDocument Map(OfficeDocumentReadResult result, string identifier, OfficeIMODocumentExtractionOptions? options) {
        ArgumentNullException.ThrowIfNull(result);
        ArgumentNullException.ThrowIfNull(identifier);
        var mapper = new ExtractionDocumentMapper(result, identifier, options ?? new OfficeIMODocumentExtractionOptions());
        return mapper.Build();
    }

    private ExtractionDocument Build() {
        AddMetadata();
        AddPageFrames();
        IndexOcrText();

        List<OfficeDocumentContentItem> content = _result.EnumerateContent().ToList();
        IndexTables(content);
        List<OfficeDocumentAsset>?[] imagesAfter = PlaceImages(content, out List<OfficeDocumentAsset> unplacedImages);
        for (int i = 0; i < content.Count; i++) {
            AddContent(content[i]);
            if (imagesAfter[i] is { } images) {
                foreach (OfficeDocumentAsset asset in images) AddImage(asset, CurrentSection());
            }
        }
        if (unplacedImages.Count > 0) {
            // Images without a source position must not look as if they belong to the last heading.
            var trailing = new ExtractionDocumentSection(null);
            _document.Sections.Add(trailing);
            foreach (OfficeDocumentAsset asset in unplacedImages) AddImage(asset, trailing);
        }
        return _document;
    }

    private void AddMetadata() {
        Dictionary<string, object?> metadata = _document.Metadata;
        OfficeDocumentSource source = _result.Source;
        SetIfPresent(metadata, "title", source.Title);
        SetIfPresent(metadata, "author", source.Author);
        SetIfPresent(metadata, "subject", source.Subject);
        SetIfPresent(metadata, "keywords", source.Keywords);
        metadata[OfficeIMOExtractionPropertyNames.InputKind] = _result.Kind.ToString();
        SetIfPresent(metadata, OfficeIMOExtractionPropertyNames.SourceHash, source.SourceHash);

        foreach (OfficeDocumentMetadataEntry entry in _result.Metadata) {
            if (string.IsNullOrWhiteSpace(entry.Name) || entry.Value == null) continue;
            string key = string.IsNullOrWhiteSpace(entry.Category) ? entry.Name : entry.Category + "." + entry.Name;
            if (!metadata.ContainsKey(key)) metadata[key] = entry.Value;
        }

        List<Dictionary<string, object?>> links = _result.Links.Select(DescribeLink).ToList();
        if (links.Count > 0) metadata[OfficeIMOExtractionPropertyNames.Links] = links;
        List<Dictionary<string, object?>> forms = _result.Forms.Select(DescribeFormField).ToList();
        if (forms.Count > 0) metadata[OfficeIMOExtractionPropertyNames.FormFields] = forms;
        List<Dictionary<string, object?>> diagnostics = _result.Diagnostics.Select(DescribeDiagnostic).ToList();
        if (diagnostics.Count > 0) metadata[OfficeIMOExtractionPropertyNames.Diagnostics] = diagnostics;
    }

    private void AddPageFrames() {
        // Pages are added up front, in source order, so blank pages remain visible as frames.
        // Slides and sheets are pages too; keying them by their own location gives each one frame.
        foreach (OfficeDocumentPage page in _result.Pages) {
            string? key = FrameKey(page.Location, out int? number);
            if (key == null && page.Number is int pageNumber && pageNumber > 0) {
                key = PageKey(pageNumber);
                number = pageNumber;
            }
            if (key == null || _pagesByKey.ContainsKey(key)) continue;
            _pagesByKey[key] = page;
            GetOrAddFrame(key, number);
        }
    }

    private void IndexTables(List<OfficeDocumentContentItem> content) {
        HashSet<string> anchoredSignatures = new(StringComparer.Ordinal);
        foreach (OfficeDocumentContentItem item in content) {
            if (item.Table != null && ItemLocation(item)?.BlockAnchor is { Length: > 0 } anchor) {
                _tableAnchors.Add(anchor);
                anchoredSignatures.Add(TableSignature(item.Table));
            }
        }
        // A table can reach the traversal twice: from the document and as an unanchored projection of a text chunk.
        // Keep the anchored one.
        foreach (OfficeDocumentContentItem item in content) {
            if (item.Table != null && string.IsNullOrEmpty(ItemLocation(item)?.BlockAnchor) && anchoredSignatures.Contains(TableSignature(item.Table))) {
                _duplicateTables.Add(item.Table);
            }
        }
    }

    private static string TableSignature(ReaderTable table) {
        var builder = new StringBuilder();
        foreach (string column in table.Columns) builder.Append(column).Append('\u001f');
        foreach (IReadOnlyList<string> row in table.Rows) {
            builder.Append('\u001e');
            if (row == null) continue;
            foreach (string cell in row) builder.Append(cell).Append('\u001f');
        }
        return builder.ToString();
    }

    private void IndexOcrText() {
        if (_result.OcrCandidates.Count == 0) return;
        Dictionary<string, OfficeDocumentBlock> recognized = new(StringComparer.Ordinal);
        foreach (OfficeDocumentBlock block in _result.EnumerateBlocks()) {
            if (block.Recognition != null && !string.IsNullOrEmpty(block.Id) && !recognized.ContainsKey(block.Id)) recognized[block.Id] = block;
        }
        foreach (OfficeDocumentOcrCandidate candidate in _result.OcrCandidates) {
            if (string.IsNullOrEmpty(candidate.AssetId) || string.IsNullOrEmpty(candidate.Id)) continue;
            if (recognized.TryGetValue(candidate.Id + "-block", out OfficeDocumentBlock? block) && !_ocrTextByAssetId.ContainsKey(candidate.AssetId!)) {
                _ocrTextByAssetId[candidate.AssetId!] = block;
            }
        }
    }

    private void AddContent(OfficeDocumentContentItem item) {
        if (item.Table != null) {
            AddTable(item.Table, item.Location);
        } else if (item.Block != null) {
            AddBlock(item.Block, item.Location ?? item.Block.Location);
        } else if (item.Chunk != null && !string.IsNullOrWhiteSpace(item.Chunk.Text)) {
            var payload = new TextContent(item.Chunk.Text) { RawRepresentation = item.Chunk };
            var element = new ExtractionDocumentTextElement(payload, null);
            AddLocationProperties(payload, item.Location ?? item.Chunk.Location);
            AddSources(element.Sources, item.Location ?? item.Chunk.Location, null);
            CurrentSection().Elements.Add(element);
        }
    }

    private void AddBlock(OfficeDocumentBlock block, ReaderLocation location) {
        // OCR text that belongs to an image is written to that image element instead.
        if (!string.IsNullOrEmpty(block.Id) && _consumedOcrBlockIds.Contains(block.Id)) return;
        string kind = block.Kind ?? string.Empty;
        if (_imagesForBlocks.TryGetValue(block, out OfficeDocumentAsset? pairedImage)) {
            AddImage(pairedImage, CurrentSection(), location, block.Text);
            return;
        }
        if (string.IsNullOrWhiteSpace(block.Text)) return;
        // A block that shares a table's anchor is that table's text projection (a Word table block, a worksheet
        // block); the structured table is emitted instead.
        if (!string.IsNullOrEmpty(location.BlockAnchor) && _tableAnchors.Contains(location.BlockAnchor!)) return;
        if (string.Equals(kind, "table", StringComparison.OrdinalIgnoreCase)) return;

        var payload = new TextContent(block.Text) { RawRepresentation = block };
        string? role = MapRole(kind);
        var element = new ExtractionDocumentTextElement(payload, role);
        AddProperty(payload, OfficeIMOExtractionPropertyNames.BlockKind, kind);
        if (!string.IsNullOrEmpty(block.Id)) AddProperty(payload, OfficeIMOExtractionPropertyNames.BlockId, block.Id);
        AddLocationProperties(payload, location);
        AddSources(element.Sources, location, block.Region);

        if (role == "heading") {
            int level = block.Level is int value && value > 0 ? value : 1;
            element.HeadingLevel = level;
            StartSection(level).Elements.Add(element);
            return;
        }

        if (role == "listItem") {
            if (block.Level is int listLevel) AddProperty(payload, OfficeIMOExtractionPropertyNames.ListLevel, listLevel);
            if (!string.IsNullOrEmpty(block.Marker)) AddProperty(payload, OfficeIMOExtractionPropertyNames.ListMarker, block.Marker);
            if (block.ListIndex is int listIndex) AddProperty(payload, OfficeIMOExtractionPropertyNames.ListIndex, listIndex);
        }

        CurrentSection().Elements.Add(element);
    }

    private void AddTable(ReaderTable table, ReaderLocation? location) {
        if (_duplicateTables.Contains(table)) return;
        // Readers name columns "Column 1", "Column 2"... when the source has no header row; those are not headers.
        bool hasHeader = table.Columns.Count > 0 && !IsFallbackHeader(table.Columns);
        int headerRows = hasHeader ? 1 : 0;
        int columnCount = table.Columns.Count;
        foreach (IReadOnlyList<string> row in table.Rows) columnCount = Math.Max(columnCount, row?.Count ?? 0);
        int rowCount = headerRows + table.Rows.Count;
        if (rowCount == 0 || columnCount == 0) return;

        // A structural table has no text payload of its own; a bare AIContent carries its properties and native table.
        var payload = new AIContent { RawRepresentation = table };
        var element = new ExtractionDocumentTableElement(payload, rowCount, columnCount);
        if (!string.IsNullOrWhiteSpace(table.Title)) AddProperty(payload, OfficeIMOExtractionPropertyNames.TableTitle, table.Title);
        if (!string.IsNullOrWhiteSpace(table.Kind)) AddProperty(payload, OfficeIMOExtractionPropertyNames.TableKind, table.Kind);
        if (table.Truncated) {
            AddProperty(payload, OfficeIMOExtractionPropertyNames.TableTruncated, true);
            AddProperty(payload, OfficeIMOExtractionPropertyNames.TableTotalRowCount, table.TotalRowCount);
        }
        ReaderLocation effective = location ?? table.Location ?? new ReaderLocation();
        AddLocationProperties(payload, effective);
        AddSources(element.Sources, effective, TableRegion(table));

        for (int column = 0; hasHeader && column < table.Columns.Count; column++) {
            string name = table.Columns[column];
            element.Cells[0, column] = CreateCell(IsFallbackColumnName(name, column) ? null : name, ColumnHeaderRole);
        }
        for (int row = 0; row < table.Rows.Count; row++) {
            IReadOnlyList<string>? cells = table.Rows[row];
            if (cells == null) continue;
            for (int column = 0; column < cells.Count; column++) {
                element.Cells[headerRows + row, column] = CreateCell(cells[column], null);
            }
        }

        CurrentSection().Elements.Add(element);
    }

    private static bool IsFallbackHeader(IReadOnlyList<string> columns) {
        for (int i = 0; i < columns.Count; i++) {
            if (!IsFallbackColumnName(columns[i], i)) return false;
        }
        return true;
    }

    private static bool IsFallbackColumnName(string? name, int index) =>
        string.Equals(name, "Column " + (index + 1).ToString(CultureInfo.InvariantCulture), StringComparison.Ordinal);

    private static ExtractionDocumentTableCell CreateCell(string? text, string? role) {
        if (string.IsNullOrEmpty(text)) return new ExtractionDocumentTableCell(1, 1, role);
        return new ExtractionDocumentTableCell(new ExtractionDocumentTextElement(new TextContent(text), null), 1, 1, role);
    }

    private void AddImage(OfficeDocumentAsset asset, ExtractionDocumentSection section, ReaderLocation? position = null, string? blockText = null) {
        AIContent payload = _options.IncludeImagePayloads && asset.PayloadBytes is { Length: > 0 } bytes && !string.IsNullOrWhiteSpace(asset.MediaType)
            ? new DataContent(bytes, asset.MediaType!) { Name = asset.FileName }
            : new AIContent();
        payload.RawRepresentation = asset;

        string? ocrText = null;
        if (!string.IsNullOrEmpty(asset.Id) && _ocrTextByAssetId.TryGetValue(asset.Id, out OfficeDocumentBlock? ocrBlock)) {
            ocrText = ocrBlock.Text;
            if (!string.IsNullOrEmpty(ocrBlock.Recognition?.Provider)) AddProperty(payload, OfficeIMOExtractionPropertyNames.OcrProvider, ocrBlock.Recognition!.Provider);
            if (ocrBlock.Recognition?.Confidence is double confidence) AddProperty(payload, OfficeIMOExtractionPropertyNames.OcrConfidence, confidence);
        }

        string? alternativeText = !string.IsNullOrWhiteSpace(asset.AltText) ? asset.AltText
            : !string.IsNullOrWhiteSpace(blockText) ? blockText : null;
        var element = new ExtractionDocumentImageElement(payload, ocrText) { AlternativeText = alternativeText };
        if (!string.IsNullOrEmpty(asset.Id)) AddProperty(payload, OfficeIMOExtractionPropertyNames.AssetId, asset.Id);
        if (!string.IsNullOrWhiteSpace(asset.Title)) AddProperty(payload, OfficeIMOExtractionPropertyNames.AssetTitle, asset.Title);
        if (!string.IsNullOrWhiteSpace(asset.FileName)) AddProperty(payload, OfficeIMOExtractionPropertyNames.AssetFileName, asset.FileName);
        if (asset.Width is int width) AddProperty(payload, OfficeIMOExtractionPropertyNames.ImageWidth, width);
        if (asset.Height is int height) AddProperty(payload, OfficeIMOExtractionPropertyNames.ImageHeight, height);
        AddLocationProperties(payload, asset.Location);
        AddSources(element.Sources, FrameKey(asset.Location, out _) != null || position == null ? asset.Location : position, asset.Region);
        section.Elements.Add(element);
    }

    /// <summary>
    /// Assets are not part of the content traversal. Each image goes after the content item that shares its anchor,
    /// otherwise after the last item on the same page, slide or sheet. The rest are returned as unplaced.
    /// </summary>
    private List<OfficeDocumentAsset>?[] PlaceImages(List<OfficeDocumentContentItem> content, out List<OfficeDocumentAsset> trailing) {
        var placement = new List<OfficeDocumentAsset>?[content.Count];
        trailing = new List<OfficeDocumentAsset>();
        if (_result.Assets.Count == 0) return placement;

        Dictionary<string, int> byAnchor = new(StringComparer.Ordinal);
        Dictionary<string, int> lastByFrame = new(StringComparer.Ordinal);
        for (int i = 0; i < content.Count; i++) {
            ReaderLocation? location = ItemLocation(content[i]);
            if (location == null) continue;
            if (!string.IsNullOrEmpty(location.BlockAnchor)) byAnchor[location.BlockAnchor!] = i;
            string? frame = FrameKey(location, out _);
            if (frame != null) lastByFrame[frame] = i;
        }

        List<OfficeDocumentAsset> images = _result.Assets.Where(IsImage).ToList();
        List<OfficeDocumentBlock> imageBlocks = content
            .Where(item => item.Block != null && string.Equals(item.Block.Kind, "image", StringComparison.OrdinalIgnoreCase))
            .Select(item => item.Block!)
            .ToList();
        // Some readers (HTML, EPUB) emit an image block in document order plus a separately numbered asset. When the
        // counts agree, the n-th block is the n-th asset, which gives the image its position.
        bool pairByOrder = imageBlocks.Count > 0 && imageBlocks.Count == images.Count;
        for (int i = 0; pairByOrder && i < images.Count; i++) _imagesForBlocks[imageBlocks[i]] = images[i];

        foreach (OfficeDocumentAsset asset in images) {
            if (!string.IsNullOrEmpty(asset.Id) && _ocrTextByAssetId.TryGetValue(asset.Id, out OfficeDocumentBlock? ocrBlock)) {
                _consumedOcrBlockIds.Add(ocrBlock.Id);
            }
            if (pairByOrder) continue;
            int index = -1;
            ReaderLocation location = asset.Location;
            if (!string.IsNullOrEmpty(location.BlockAnchor) && byAnchor.TryGetValue(location.BlockAnchor!, out int anchored)) {
                index = anchored;
            } else if (FrameKey(location, out _) is { } frame && lastByFrame.TryGetValue(frame, out int last)) {
                index = last;
            }
            if (index < 0) {
                trailing.Add(asset);
            } else {
                (placement[index] ??= new List<OfficeDocumentAsset>()).Add(asset);
            }
        }
        return placement;
    }

    private static ReaderLocation? ItemLocation(OfficeDocumentContentItem item) =>
        item.Location ?? item.Table?.Location ?? item.Block?.Location ?? item.Chunk?.Location;

    private static bool IsImage(OfficeDocumentAsset asset) =>
        string.Equals(asset.Kind, "image", StringComparison.OrdinalIgnoreCase)
        || (asset.MediaType?.StartsWith("image/", StringComparison.OrdinalIgnoreCase) ?? false);

    private ExtractionDocumentSection StartSection(int level) {
        while (_headings.Count > 0 && _headings[_headings.Count - 1].Level >= level) _headings.RemoveAt(_headings.Count - 1);
        var section = new ExtractionDocumentSection(null);
        if (_headings.Count == 0) {
            _document.Sections.Add(section);
        } else {
            _headings[_headings.Count - 1].Section.Elements.Add(section);
        }
        _headings.Add((level, section));
        return section;
    }

    private ExtractionDocumentSection CurrentSection() {
        if (_headings.Count > 0) return _headings[_headings.Count - 1].Section;
        if (_preamble == null) {
            // Content before the first heading still needs a root section, because documents hold sections only.
            _preamble = new ExtractionDocumentSection(null);
            _document.Sections.Add(_preamble);
        }
        return _preamble;
    }

    private void AddSources(List<ExtractionDocumentSourceReference> sources, ReaderLocation? location, OfficeDocumentRegion? region) {
        if (location == null) return;
        string? key = FrameKey(location, out int? pageNumber);
        if (key == null) return;
        ExtractionDocumentSourcePage frame = GetOrAddFrame(key, pageNumber);
        IEnumerable<ExtractionDocumentPoint> polygon = BuildPolygon(region);
        sources.Add(new ExtractionDocumentSourceReference(frame, polygon, null));
    }

    private IEnumerable<ExtractionDocumentPoint> BuildPolygon(OfficeDocumentRegion? region) {
        // Only PDF reports regions in a known frame (points, bottom-left origin). Text lines are baselines with no
        // height; a page-only reference is more honest than a zero-area polygon.
        if (!_options.IncludeRegions || !_hasPdfGeometry || region == null) return Array.Empty<ExtractionDocumentPoint>();
        if (!(region.Width > 0) || !(region.Height > 0)) return Array.Empty<ExtractionDocumentPoint>();
        double left = region.X, bottom = region.Y, right = region.X + region.Width, top = region.Y + region.Height;
        return new[] {
            new ExtractionDocumentPoint(left, top),
            new ExtractionDocumentPoint(right, top),
            new ExtractionDocumentPoint(right, bottom),
            new ExtractionDocumentPoint(left, bottom)
        };
    }

    private static OfficeDocumentRegion? TableRegion(ReaderTable table) {
        ReaderTableDiagnostics? diagnostics = table.Diagnostics;
        if (diagnostics == null || !diagnostics.HasGeometry) return null;
        return new OfficeDocumentRegion {
            X = diagnostics.XStart,
            Y = Math.Min(diagnostics.YBottom, diagnostics.YTop),
            Width = diagnostics.XEnd - diagnostics.XStart,
            Height = Math.Abs(diagnostics.YTop - diagnostics.YBottom)
        };
    }

    private ExtractionDocumentSourcePage GetOrAddFrame(string key, int? pageNumber) {
        if (_frames.TryGetValue(key, out ExtractionDocumentSourcePage? existing)) return existing;
        ExtractionDocumentPageDimensions? dimensions = null;
        ExtractionDocumentCoordinateUnit? unit = null;
        ExtractionDocumentCoordinateOrigin? origin = null;
        if (_hasPdfGeometry && _pagesByKey.TryGetValue(key, out OfficeDocumentPage? page)
            && page.Width is double width && page.Height is double height && width > 0 && height > 0) {
            dimensions = new ExtractionDocumentPageDimensions(width, height);
            unit = ExtractionDocumentCoordinateUnit.Point;
            origin = ExtractionDocumentCoordinateOrigin.BottomLeft;
        }
        var frame = new ExtractionDocumentSourcePage(key, pageNumber, dimensions, unit, origin);
        _frames[key] = frame;
        _document.SourcePages.Add(frame);
        return frame;
    }

    private static string? FrameKey(ReaderLocation location, out int? pageNumber) {
        pageNumber = null;
        if (location.Page is int page && page > 0) {
            pageNumber = page;
            return PageKey(page);
        }
        if (location.Slide is int slide && slide > 0) {
            pageNumber = slide;
            return "slide-" + slide.ToString(CultureInfo.InvariantCulture);
        }
        if (!string.IsNullOrEmpty(location.Sheet)) return "sheet-" + location.Sheet;
        return null;
    }

    private static string PageKey(int number) => "page-" + number.ToString(CultureInfo.InvariantCulture);

    private static void AddLocationProperties(AIContent payload, ReaderLocation? location) {
        if (location == null) return;
        if (!string.IsNullOrEmpty(location.Sheet)) AddProperty(payload, OfficeIMOExtractionPropertyNames.Sheet, location.Sheet);
        if (!string.IsNullOrEmpty(location.A1Range)) AddProperty(payload, OfficeIMOExtractionPropertyNames.CellRange, location.A1Range);
        if (location.Slide is int slide) AddProperty(payload, OfficeIMOExtractionPropertyNames.Slide, slide);
        if (!string.IsNullOrEmpty(location.HeadingPath)) AddProperty(payload, OfficeIMOExtractionPropertyNames.HeadingPath, location.HeadingPath);
        if (!string.IsNullOrEmpty(location.BlockAnchor)) AddProperty(payload, OfficeIMOExtractionPropertyNames.BlockAnchor, location.BlockAnchor);
    }

    private static void AddProperty(AIContent payload, string key, object? value) {
        (payload.AdditionalProperties ??= new AdditionalPropertiesDictionary())[key] = value;
    }

    private static string? MapRole(string kind) {
        switch (kind.ToLowerInvariant()) {
            case "heading":
                return "heading";
            case "paragraph":
            case "text":
            case "text-block":
            case "":
                return null;
            case "list-item":
                return "listItem";
            case "header":
            case "page-header":
                return "pageHeader";
            case "footer":
            case "page-footer":
                return "pageFooter";
            default:
                // Custom roles are preserved by the shared model, so footnotes, comments, speaker notes and the like
                // keep the OfficeIMO kind as their role.
                return kind;
        }
    }

    private static Dictionary<string, object?> DescribeLink(OfficeDocumentLink link) {
        var value = new Dictionary<string, object?>(StringComparer.Ordinal) { ["kind"] = link.Kind };
        SetIfPresent(value, "text", link.Text);
        SetIfPresent(value, "uri", link.Uri);
        SetIfPresent(value, "destinationName", link.DestinationName);
        if (link.DestinationPageNumber is int destinationPage) value["destinationPage"] = destinationPage;
        if (link.Location.Page is int page) value["page"] = page;
        SetIfPresent(value, "blockAnchor", link.Location.BlockAnchor);
        return value;
    }

    private static Dictionary<string, object?> DescribeFormField(OfficeDocumentFormField field) {
        var value = new Dictionary<string, object?>(StringComparer.Ordinal) { ["kind"] = field.Kind };
        SetIfPresent(value, "name", field.Name);
        SetIfPresent(value, "value", field.Value);
        value["readOnly"] = field.IsReadOnly;
        value["required"] = field.IsRequired;
        if (field.Location.Page is int page) value["page"] = page;
        return value;
    }

    private static Dictionary<string, object?> DescribeDiagnostic(OfficeDocumentDiagnostic diagnostic) => new(StringComparer.Ordinal) {
        ["severity"] = diagnostic.Severity.ToString(),
        ["category"] = diagnostic.Category.ToString(),
        ["code"] = diagnostic.Code,
        ["message"] = diagnostic.Message
    };

    private static void SetIfPresent(Dictionary<string, object?> target, string key, string? value) {
        if (!string.IsNullOrWhiteSpace(value)) target[key] = value;
    }
}
