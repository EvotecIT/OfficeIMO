using System.Security.Cryptography;

namespace OfficeIMO.Xps;

public sealed partial class XpsDocument {
    /// <summary>Projects literal Unicode, native reading order and semantic structure into the shared document model.</summary>
    /// <param name="sourceName">Optional source path or logical name.</param>
    /// <param name="includeSvgPreviewAssets">Includes self-contained page SVG assets through the strict native renderer.</param>
    /// <param name="cancellationToken">Cancels native structure reading and rendering.</param>
    /// <remarks>Unreferenced Unicode follows declared stories in page/markup order. Glyph identifiers are not reverse-mapped to text. Cell spans remain in the recursive structure; rectangular table projections use blank occupied cells. Preview assets are limited to 512 pages and 128 MiB in aggregate.</remarks>
    public OfficeDocumentModel ToOfficeDocumentModel(string? sourceName = null, bool includeSvgPreviewAssets = false,
        CancellationToken cancellationToken = default) => new XpsDocumentModelProjection(this, sourceName, cancellationToken).Build(includeSvgPreviewAssets);
}

internal sealed partial class XpsDocumentModelProjection {
    private readonly XpsDocument _document;
    private readonly string? _source;
    private readonly CancellationToken _token;
    private readonly XpsStoryFragmentsReader.Budget _budget;
    private readonly List<OfficeDocumentModelBlock> _blocks = new();
    private readonly List<OfficeDocumentModelTable> _tables = new();
    private readonly List<OfficeDocumentModelDiagnostic> _diagnostics = new();
    private readonly List<OfficeDocumentModelLink> _links = new();
    private readonly Dictionary<OfficeDocumentModelTable, HashSet<int>> _tablePages = new();
    private readonly HashSet<(int Page, int Glyph)> _referenced = new();
    private int _nodeIndex;
    private long _logicalOrder;

    internal XpsDocumentModelProjection(XpsDocument document, string? sourceName, CancellationToken token) {
        _document = document; _source = sourceName; _token = token; _budget = new XpsStoryFragmentsReader.Budget(token);
    }

    internal OfficeDocumentModel Build(bool previews) {
        var logical = _document.ReadLogicalStructure(_token);
        foreach (string message in logical.Diagnostics) Diagnostic("XpsLogicalStructure", message);
        var structure = new List<OfficeDocumentModelNode>();
        foreach (var story in logical.Stories) {
            _budget.Charge();
            string kind = story.Type == XpsStoryFragmentType.Content ? "story" : story.Type.ToString().ToLowerInvariant();
            var children = story.Blocks.Select(node => Node(node, kind, "named-element")).ToArray();
            structure.Add(new OfficeDocumentModelNode { Id = NextId(), Kind = kind, Children = children,
                Attributes = new Dictionary<string, string> { ["storyName"] = story.Name ?? string.Empty } });
        }
        for (int pageIndex = 0; pageIndex < _document.Pages.Count; pageIndex++) {
            _budget.Charge(); int ordinal = 0;
            var page = _document.Pages[pageIndex];
            var markup = page.GetMarkup(); AddLinks(page, pageIndex, markup);
            foreach (var glyph in XpsStoryFragmentsReader.PageElements(markup).Where(e => e.Name.LocalName == "Glyphs")) {
                _budget.Charge(); int current = ordinal++;
                if (_referenced.Contains((pageIndex, current))) continue;
                if (glyph.Attribute("UnicodeString") is not XAttribute unicode) {
                    Diagnostic("XpsGlyphTextUnavailable", "A native glyph run has no UnicodeString.", pageIndex); continue;
                }
                string text = XpsPage.Unescape(unicode.Value); _budget.Text(text.Length);
                AddBlock(text, "unstructured-text", pageIndex, current);
            }
            if (!logical.Pages[pageIndex].HasNativeStructure)
                Diagnostic("XpsMarkupTextOrder", "This page has no native StoryFragments; its Unicode is projected in markup order.", pageIndex);
        }
        var assets = new List<OfficeDocumentModelAsset>(); long assetBytes = 0;
        var blocksByPage = _blocks.ToLookup(block => block.Location.Page);
        var tablesByPage = _tables.Where(table => _tablePages[table].Count == 1).ToLookup(table => table.Location!.Page);
        var linksByPage = _links.ToLookup(link => link.Location.Page);
        var pages = new List<OfficeDocumentModelPage>();
        for (int pageIndex = 0; pageIndex < _document.Pages.Count; pageIndex++) {
            _budget.Charge(); var page = _document.Pages[pageIndex];
            var blocks = blocksByPage[pageIndex + 1].ToArray();
            var pageAssets = new List<OfficeDocumentModelAsset>();
            if (previews) {
                if (assets.Count >= 512) throw new InvalidDataException("XPS document projection preview count exceeds 512.");
                byte[] svg = Encoding.UTF8.GetBytes(page.ToSvg(cancellationToken: _token).Svg);
                if (svg.LongLength > 128L * 1024 * 1024 - assetBytes) throw new InvalidDataException("XPS document projection preview bytes exceed 128 MiB.");
                assetBytes += svg.LongLength;
                var asset = new OfficeDocumentModelAsset { Id = "xps-page-" + (pageIndex + 1) + "-svg", Kind = "page-preview",
                    MediaType = "image/svg+xml", Extension = ".svg", PayloadBytes = svg, LengthBytes = svg.LongLength,
                    PayloadHash = Hash(svg), SourceObjectId = page.PartName,
                    Location = Location(pageIndex, null, "page-preview") };
                assets.Add(asset); pageAssets.Add(asset);
            }
            pages.Add(new OfficeDocumentModelPage { Number = pageIndex + 1, Name = page.PartName,
                Width = page.Width * .75, Height = page.Height * .75, Text = string.Join("\n", blocks.Select(b => b.Text)),
                Blocks = blocks, Tables = tablesByPage[pageIndex + 1].ToArray(), Assets = pageAssets,
                Links = linksByPage[pageIndex + 1].ToArray(),
                Location = Location(pageIndex, null, "fixed-page") });
        }
        return new OfficeDocumentModel { Format = OfficeDocumentFormat.Xps, Source = new OfficeDocumentModelSource { Path = _source },
            CapabilitiesUsed = previews ? new[] { "officeimo.xps.unicode", "officeimo.xps.native-structure", "officeimo.xps.svg" }
                : new[] { "officeimo.xps.unicode", "officeimo.xps.native-structure" },
            Blocks = _blocks, Structure = structure, Pages = pages, Tables = _tables, Assets = assets, Links = _links, Diagnostics = _diagnostics };
    }

    private OfficeDocumentModelNode Node(XpsStructureNode node, string storyKind, string blockKind) {
        _budget.Charge(); long order = _logicalOrder++; string kind = Kind(node.Kind);
        if (node.Kind is XpsStructureKind.Paragraph or XpsStructureKind.Figure) blockKind = kind;
        var children = new List<OfficeDocumentModelNode>();
        if (node.Marker != null) children.Add(Content(node.Marker, "list-marker", storyKind));
        if (node.Content != null) return Content(node.Content, blockKind, storyKind);
        foreach (var child in node.Children) children.Add(Node(child, storyKind, blockKind));
        var location = children.Select(c => c.Location).FirstOrDefault(l => l.Page.HasValue) ?? new OfficeDocumentModelLocation { Path = _source };
        if (node.Kind == XpsStructureKind.Table) {
            var tableLocation = new OfficeDocumentModelLocation { Path = location.Path, Page = location.Page,
                LogicalOrder = order, SourceBlockKind = "table", TableIndex = _tables.Count };
            if (!AddTable(node, tableLocation)) tableLocation.TableIndex = null;
            location = tableLocation;
        }
        var attributes = node.Kind == XpsStructureKind.TableCell ? new Dictionary<string, string> {
                ["rowSpan"] = node.RowSpan.ToString(CultureInfo.InvariantCulture), ["columnSpan"] = node.ColumnSpan.ToString(CultureInfo.InvariantCulture)
            } : new Dictionary<string, string>();
        if (node.NameReference != null) attributes["nameReference"] = node.NameReference;
        return new OfficeDocumentModelNode { Id = NextId(), Kind = kind, Children = children, Location = location, Attributes = attributes };
    }

    private OfficeDocumentModelNode Content(XpsNamedContent content, string kind, string storyKind) {
        _budget.Charge();
        foreach (int ordinal in content.GlyphOrdinals) {
            if (!_referenced.Add((content.PageIndex, ordinal)))
                throw new NotSupportedException("Native XPS structure assigns the same glyph run to multiple reading-order positions.");
        }
        _budget.Text(content.Text.Length);
        string id = NextId(); int? firstOrdinal = content.GlyphOrdinals.Count == 0 ? null : content.GlyphOrdinals[0];
        var location = Location(content.PageIndex, firstOrdinal, kind);
        if (content.Text.Length != 0) location = AddBlock(content.Text, storyKind is "header" or "footer" ? storyKind : kind,
            content.PageIndex, firstOrdinal!.Value).Location;
        return new OfficeDocumentModelNode { Id = id, Kind = kind == "list-marker" ? kind : "named-element", Text = content.Text, Location = location,
            Attributes = new Dictionary<string, string> { ["nameReference"] = content.Name, ["elementName"] = content.ElementName,
                ["pagePartName"] = content.PagePartName } };
    }

    private OfficeDocumentModelBlock AddBlock(string text, string kind, int page, int ordinal) {
        var location = Location(page, ordinal, kind); location.BlockIndex = _blocks.Count; location.LogicalOrder = _logicalOrder++;
        location.BlockAnchor = "xps-block-" + _blocks.Count.ToString(CultureInfo.InvariantCulture);
        var block = new OfficeDocumentModelBlock { Id = location.BlockAnchor, Kind = kind, Text = text, Location = location };
        _blocks.Add(block); return block;
    }
    private OfficeDocumentModelLocation Location(int page, int? ordinal, string kind) => new() {
        Path = _source, Page = page + 1, SourceBlockIndex = ordinal, SourceBlockKind = kind
    };
    private string NextId() => "xps-node-" + (++_nodeIndex).ToString(CultureInfo.InvariantCulture);
    private static string Hash(byte[] payload) {
        using var sha = SHA256.Create();
        return BitConverter.ToString(sha.ComputeHash(payload)).Replace("-", string.Empty).ToLowerInvariant();
    }
    private void Diagnostic(string code, string message, int? page = null) => _diagnostics.Add(new OfficeDocumentModelDiagnostic {
        Code = code, Message = message, Source = "OfficeIMO.Xps", Category = OfficeDocumentModelDiagnosticCategory.Content,
        IsRecoverable = true, Location = page.HasValue ? Location(page.Value, null, "fixed-page") : null
    });
    private static string Kind(XpsStructureKind kind) => kind switch {
        XpsStructureKind.Section => "section", XpsStructureKind.Paragraph => "paragraph", XpsStructureKind.Table => "table",
        XpsStructureKind.TableRowGroup => "table-row-group", XpsStructureKind.TableRow => "table-row", XpsStructureKind.TableCell => "table-cell",
        XpsStructureKind.List => "list", XpsStructureKind.ListItem => "list-item", XpsStructureKind.Figure => "figure",
        XpsStructureKind.NamedElement => "named-element", _ => "unknown"
    };
}
