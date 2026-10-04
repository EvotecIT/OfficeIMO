using OfficeIMO.Pdf;

namespace OfficeIMO.Xps;

/// <summary>Maps authored native structure without laying out or replaying page paint.</summary>
internal sealed class XpsPdfLogicalStructure {
    private sealed class TextRoute {
        internal TextRoute(PdfCanvasSourceContent content, bool artifact) { Content = content; Artifact = artifact; }
        internal PdfCanvasSourceContent Content { get; }
        internal bool Artifact { get; }
    }
    private sealed class PagePlan {
        internal readonly Dictionary<int, TextRoute> Glyphs = new();
        internal readonly Dictionary<string, PdfCanvasSourceContent> Paint = new(StringComparer.Ordinal);
        internal readonly List<IReadOnlyList<PdfCanvasStructureStep>> Containers = new();
    }
    private readonly PagePlan[] _pages;
    private readonly CancellationToken _token;
    private long _order;

    internal XpsPdfLogicalStructure(XpsDocument document, CancellationToken token) {
        _token = token;
        var logical = document.ReadLogicalStructure(token);
        HasNativeStructure = logical.HasNativeStructure;
        _pages = Enumerable.Range(0, document.Pages.Count).Select(_ => new PagePlan()).ToArray();
        if (!HasNativeStructure) return;
        // Page-local fragments remain useful when no cross-page association was
        // authored. Every other loss must be explicit rather than silently tagged.
        var losses = logical.Diagnostics.Where(d => !d.StartsWith("Story fragment has no declared story address: ", StringComparison.Ordinal)).ToArray();
        if (losses.Length != 0 || logical.Pages.Any(p => p.HasUnresolvedNames || !p.IsComplete))
            throw new NotSupportedException("Native XPS structure cannot be mapped without loss. Use preserveLogicalStructure: false for paint and markup-order text. " + string.Join("; ", losses.Take(8)));
        foreach (var story in logical.Stories) {
            _token.ThrowIfCancellationRequested();
            bool artifact = story.Type != XpsStoryFragmentType.Content;
            var path = new List<PdfCanvasStructureStep>();
            if (!artifact) path.Add(Step(PdfCanvasStructureRole.Section));
            if (!artifact) _pages[story.Fragments[0].PageIndex].Containers.Add(path.ToArray());
            foreach (var block in story.Blocks) Visit(block, path, artifact, story.Fragments[0].PageIndex);
        }
    }
    internal bool HasNativeStructure { get; }
    internal IReadOnlyDictionary<string, PdfCanvasSourceContent> Paint(int page) => _pages[page].Paint;

    private PdfCanvasStructureStep Step(PdfCanvasStructureRole role, int rowSpan = 1, int columnSpan = 1) {
        long order = _order++;
        return new PdfCanvasStructureStep(role, new PdfCanvasStructureOptions {
            StructureElementKey = "xps-structure-" + order.ToString(System.Globalization.CultureInfo.InvariantCulture),
            LogicalOrder = order, RowSpan = rowSpan, ColumnSpan = columnSpan
        });
    }
    private void Visit(XpsStructureNode node, List<PdfCanvasStructureStep> path, bool artifact, int pageIndex) {
        _token.ThrowIfCancellationRequested();
        if (node.Kind == XpsStructureKind.NamedElement) {
            Add(node.Content ?? throw new NotSupportedException("Unresolved native name reference."), path, artifact);
            return;
        }
        var nested = new List<PdfCanvasStructureStep>(path);
        var role = node.Kind switch {
            XpsStructureKind.Section => PdfCanvasStructureRole.Section,
            XpsStructureKind.Paragraph => PdfCanvasStructureRole.Paragraph,
            XpsStructureKind.Table => PdfCanvasStructureRole.Table,
            XpsStructureKind.TableRowGroup => PdfCanvasStructureRole.TableBody,
            XpsStructureKind.TableRow => PdfCanvasStructureRole.TableRow,
            XpsStructureKind.TableCell => PdfCanvasStructureRole.TableCell,
            XpsStructureKind.List => PdfCanvasStructureRole.List,
            XpsStructureKind.ListItem => PdfCanvasStructureRole.ListItem,
            XpsStructureKind.Figure => PdfCanvasStructureRole.Figure,
            _ => throw new NotSupportedException("Unsupported native structure role: " + node.Kind)
        };
        nested.Add(Step(role, node.RowSpan, node.ColumnSpan));
        if (node.Kind == XpsStructureKind.ListItem) {
            if (node.Marker != null) {
                var label = new List<PdfCanvasStructureStep>(nested) { Step(PdfCanvasStructureRole.ListLabel) };
                if (!artifact) _pages[node.Marker.PageIndex].Containers.Add(label.ToArray());
                Add(node.Marker, label, artifact);
            }
            nested.Add(Step(PdfCanvasStructureRole.ListBody));
        }
        // Authored containers exist even when all descendants are non-emitting
        // (zero-size glyphs, empty paths, or collapsed transforms).
        if (!artifact) _pages[pageIndex].Containers.Add(nested.ToArray());
        foreach (var child in node.Children) Visit(child, nested, artifact, pageIndex);
    }
    private void Add(XpsNamedContent native, IReadOnlyList<PdfCanvasStructureStep> path, bool artifact) {
        var page = _pages[native.PageIndex];
        var content = new PdfCanvasSourceContent(path.ToArray(), _order++, artifact);
        var route = new TextRoute(content, artifact);
        foreach (int glyph in native.GlyphOrdinals) {
            _token.ThrowIfCancellationRequested();
            if (page.Glyphs.ContainsKey(glyph))
                throw new NotSupportedException("Native structure assigns the same glyph content more than once; PDF export cannot duplicate its semantic ownership.");
            page.Glyphs.Add(glyph, route);
        }
        if (native.ElementName != "Glyphs" || path.Any(step => step.Role == PdfCanvasStructureRole.Figure)) {
            if (page.Paint.ContainsKey("xps-" + native.Name))
                throw new NotSupportedException("Native structure assigns the same graphic more than once.");
            page.Paint.Add("xps-" + native.Name, content);
        }
    }

    internal void AddText(PdfPageCanvas canvas, IReadOnlyList<XpsTextSpan> spans, int pageIndex) {
        var page = _pages[pageIndex];
        foreach (var path in page.Containers) canvas.SourceStructureContainer(path);
        // Preserve every native Unicode cluster. Unreferenced text retains its
        // markup order in a generic division rather than receiving an inferred role.
        var groups = spans.GroupBy(span => page.Glyphs.TryGetValue(span.GlyphOrdinal, out var route) ? route : null)
            .OrderBy(group => group.Key?.Content.LogicalOrder ?? long.MaxValue);
        PdfCanvasStructureStep? fallback = null;
        foreach (var group in groups) {
            _token.ThrowIfCancellationRequested();
            var route = group.Key;
            IReadOnlyList<PdfCanvasStructureStep> path = route?.Content.Path ?? new[] { fallback ??= Step(PdfCanvasStructureRole.Division) };
            void AddSpans(PdfPageCanvas target) {
                foreach (var span in group) {
                    _token.ThrowIfCancellationRequested();
                    target.SearchableText(span.Text, Quad(span), route?.Content.LogicalOrder ?? fallback!.Options.LogicalOrder);
                }
            }
            if (route?.Artifact == true) canvas.Artifact(AddSpans);
            else AddPath(canvas, path, 0, AddSpans);
        }
    }
    private static void AddPath(PdfPageCanvas canvas, IReadOnlyList<PdfCanvasStructureStep> path, int index, Action<PdfPageCanvas> add) {
        if (index == path.Count) { add(canvas); return; }
        var step = path[index];
        canvas.Structure(step.Role, child => AddPath(child, path, index + 1, add), step.Options);
    }
    internal static PdfSelectionQuad Quad(XpsTextSpan span) => new(Point(span.TopLeft), Point(span.TopRight), Point(span.BottomRight), Point(span.BottomLeft));
    private static PdfSelectionPoint Point(OfficeIMO.Drawing.OfficePoint point) => new(point.X * 0.75, point.Y * 0.75);
}
