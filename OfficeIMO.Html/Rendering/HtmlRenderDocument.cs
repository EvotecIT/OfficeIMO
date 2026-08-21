using System.Collections.ObjectModel;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

/// <summary>
/// Backend-neutral result shared by HTML image and PDF export.
/// </summary>
public sealed class HtmlRenderDocument {
    private readonly ReadOnlyCollection<HtmlRenderPage> _pages;
    private readonly OfficeFontFaceCollection _fonts;
    private readonly ReadOnlyCollection<HtmlRenderHeading> _headings;
    private readonly HtmlDiagnosticReport _diagnosticReport;

    internal HtmlRenderDocument(HtmlRenderMode mode, IEnumerable<HtmlRenderPage> pages, HtmlDiagnosticReport diagnostics, OfficeFontFaceCollection? fonts = null, HtmlRenderMetadata? metadata = null, IReadOnlyDictionary<int, HtmlRenderBookmarkDefinition>? bookmarks = null) {
        Mode = mode;
        _pages = new List<HtmlRenderPage>(pages ?? throw new ArgumentNullException(nameof(pages))).AsReadOnly();
        if (_pages.Count == 0) {
            throw new ArgumentException("A rendered HTML document requires at least one surface.", nameof(pages));
        }

        _diagnosticReport = (diagnostics ?? throw new ArgumentNullException(nameof(diagnostics))).Clone();
        _fonts = fonts?.Clone() ?? new OfficeFontFaceCollection();
        Metadata = metadata ?? new HtmlRenderMetadata(null, null);
        _headings = BuildHeadings(_pages, bookmarks).AsReadOnly();
    }

    /// <summary>Layout mode used to produce the result.</summary>
    public HtmlRenderMode Mode { get; }

    /// <summary>Rendered pages, or one page for continuous output.</summary>
    public IReadOnlyList<HtmlRenderPage> Pages => _pages;

    /// <summary>Diagnostics emitted while parsing, laying out, and preparing paint operations.</summary>
    public IReadOnlyList<HtmlDiagnostic> Diagnostics => _diagnosticReport.Diagnostics;

    /// <summary>
    /// Whether rendering reported an approximation, omission, or failure. Renderer warnings are
    /// deliberately treated as loss unless they are informational diagnostics.
    /// </summary>
    public bool HasLoss => _diagnosticReport.Any(static diagnostic =>
        diagnostic.LossKind != OfficeConversionLossKind.None
        || diagnostic.Severity == HtmlDiagnosticSeverity.Warning
        || diagnostic.Severity == HtmlDiagnosticSeverity.Error);

    /// <summary>Throws with the complete structured report when the render was not lossless.</summary>
    public HtmlRenderDocument RequireNoLoss() {
        if (HasLoss) throw new HtmlConversionException(Diagnostics);
        return this;
    }

    internal HtmlDiagnosticReport DiagnosticReport => _diagnosticReport;

    /// <summary>Independent snapshot of scoped font faces retained for image and PDF backends.</summary>
    public OfficeFontFaceCollection Fonts => _fonts.Clone();

    /// <summary>Source document metadata retained for image and PDF adapters.</summary>
    public HtmlRenderMetadata Metadata { get; }

    /// <summary>Source headings retained in document order for navigation-capable backends.</summary>
    public IReadOnlyList<HtmlRenderHeading> Headings => _headings;

    /// <summary>Concatenated logical searchable text retained by the shared render model.</summary>
    public string Text => string.Join("\n", _pages.SelectMany(page => EnumerateLogicalText(page.Scene)));

    private static IEnumerable<string> EnumerateLogicalText(IEnumerable<HtmlRenderVisual> visuals) {
        foreach (HtmlRenderVisual visual in OrderForLogicalText(visuals)) {
            if (visual is HtmlRenderSemanticGroup { Role: HtmlRenderSemanticGroupRole.Artifact }) {
                continue;
            }
            if (visual is HtmlRenderLogicalTextGroup logicalTextGroup) {
                if (!ContainsArtifactVisual(logicalTextGroup.Visuals)) {
                    yield return logicalTextGroup.Text;
                } else {
                    string visibleText = string.Concat(EnumerateLogicalText(logicalTextGroup.Visuals));
                    if (visibleText.Length > 0) yield return visibleText;
                }
                continue;
            }
            if (visual is HtmlRenderText text) {
                yield return text.Text;
                continue;
            }

            IEnumerable<HtmlRenderVisual>? children = ChildVisuals(visual);
            if (children == null) continue;
            foreach (string textValue in EnumerateLogicalText(children)) yield return textValue;
        }
    }

    private static IEnumerable<HtmlRenderVisual> OrderForLogicalText(IEnumerable<HtmlRenderVisual> visuals) {
        List<HtmlRenderVisual> ordered = visuals.OrderBy(item => item.PaintOrder).ToList();
        var logicalOrders = new Dictionary<HtmlRenderVisual, int?>();
        bool hasLogicalText = false;
        bool allTextIsOrdered = true;
        foreach (HtmlRenderVisual visual in ordered) {
            int? logicalOrder = ResolveLogicalTextOrder(visual, out bool containsText);
            logicalOrders[visual] = logicalOrder;
            hasLogicalText |= containsText;
            if (containsText && !logicalOrder.HasValue) allTextIsOrdered = false;
        }
        if (!hasLogicalText || !allTextIsOrdered) return ordered;
        return ordered
            .OrderBy(visual => logicalOrders[visual] ?? int.MaxValue)
            .ThenBy(visual => visual.PaintOrder);
    }

    private static int? ResolveLogicalTextOrder(HtmlRenderVisual visual, out bool containsText) {
        if (visual is HtmlRenderSemanticGroup { Role: HtmlRenderSemanticGroupRole.Artifact }) {
            containsText = false;
            return null;
        }
        if (visual is HtmlRenderText text) {
            containsText = true;
            return text.LogicalTextOrder;
        }
        if (visual is HtmlRenderLogicalTextGroup logicalTextGroup) {
            int? logicalOrder = null;
            foreach (HtmlRenderVisual child in logicalTextGroup.Visuals) {
                int? childOrder = ResolveLogicalTextOrder(child, out _);
                if (childOrder.HasValue && (!logicalOrder.HasValue || childOrder.Value < logicalOrder.Value)) logicalOrder = childOrder;
            }
            containsText = true;
            return logicalOrder;
        }
        IEnumerable<HtmlRenderVisual>? children = ChildVisuals(visual);
        if (children == null) {
            containsText = false;
            return null;
        }
        containsText = false;
        int? order = null;
        foreach (HtmlRenderVisual child in children) {
            int? childOrder = ResolveLogicalTextOrder(child, out bool childContainsText);
            containsText |= childContainsText;
            if (childOrder.HasValue && (!order.HasValue || childOrder.Value < order.Value)) order = childOrder;
        }
        return order;
    }

    private static bool ContainsArtifactVisual(IEnumerable<HtmlRenderVisual> visuals) {
        foreach (HtmlRenderVisual visual in visuals) {
            if (visual is HtmlRenderSemanticGroup { Role: HtmlRenderSemanticGroupRole.Artifact }) return true;
            IEnumerable<HtmlRenderVisual>? children = ChildVisuals(visual);
            if (children != null && ContainsArtifactVisual(children)) return true;
        }
        return false;
    }

    private static IEnumerable<HtmlRenderVisual> EnumerateVisuals(IEnumerable<HtmlRenderVisual> visuals) {
        foreach (HtmlRenderVisual visual in visuals) {
            yield return visual;
            if (visual is HtmlRenderSemanticGroup { Role: HtmlRenderSemanticGroupRole.Artifact }) {
                continue;
            }
            IEnumerable<HtmlRenderVisual>? children = ChildVisuals(visual);
            if (children == null) continue;
            foreach (HtmlRenderVisual child in EnumerateVisuals(children)) yield return child;
        }
    }

    private static IEnumerable<HtmlRenderVisual>? ChildVisuals(HtmlRenderVisual visual) => visual is HtmlRenderClipGroup clipGroup
        ? clipGroup.Visuals
        : visual is HtmlRenderPathClipGroup pathClipGroup
            ? pathClipGroup.Visuals
        : visual is HtmlRenderEffectGroup effectGroup ? effectGroup.Visuals
        : visual is HtmlRenderSemanticGroup semanticGroup ? semanticGroup.Visuals
        : visual is HtmlRenderLayoutRegion layoutRegion ? layoutRegion.Visuals
        : visual is HtmlRenderLogicalTextGroup logicalTextGroup ? logicalTextGroup.Visuals
        : visual is HtmlRenderFormField formField ? formField.Visuals : null;

    private static List<HtmlRenderHeading> BuildHeadings(IReadOnlyList<HtmlRenderPage> pages, IReadOnlyDictionary<int, HtmlRenderBookmarkDefinition>? bookmarks) {
        var fragments = new List<(int NodeId, int Level, string Text, int PageNumber, double X, double Y, int Order, int? LogicalOrder, bool IsAnchor)>();
        foreach (HtmlRenderPage page in pages) {
            foreach (HtmlRenderBookmarkAnchor anchor in EnumerateVisuals(page.Scene).OfType<HtmlRenderBookmarkAnchor>()) {
                if (bookmarks == null || !bookmarks.TryGetValue(anchor.SemanticNodeId, out HtmlRenderBookmarkDefinition? definition) || definition.Suppressed) continue;
                fragments.Add((anchor.SemanticNodeId, definition.Level, anchor.Text, page.PageNumber, anchor.X, anchor.Y, anchor.PaintOrder, null, true));
            }
            foreach (HtmlRenderTextFragment text in EnumerateTextFragments(page.Scene)) {
                if (!text.SemanticNodeId.HasValue) continue;
                bool automatic = HtmlRenderHeading.TryGetLevel(text.SemanticRole, out int level);
                if (bookmarks != null && bookmarks.TryGetValue(text.SemanticNodeId.Value, out HtmlRenderBookmarkDefinition? definition)) {
                    if (definition.Suppressed) continue;
                    level = definition.Level;
                } else if (!automatic) continue;
                fragments.Add((text.SemanticNodeId.Value, level, text.Text, page.PageNumber, text.X, text.Y, text.PaintOrder, text.LogicalOrder, false));
            }
        }

        var headings = new List<HtmlRenderHeading>();
        foreach (IGrouping<int, (int NodeId, int Level, string Text, int PageNumber, double X, double Y, int Order, int? LogicalOrder, bool IsAnchor)> group in fragments
            .GroupBy(item => item.NodeId)
            .OrderBy(group => bookmarks != null && bookmarks.TryGetValue(group.Key, out HtmlRenderBookmarkDefinition? definition)
                ? definition.SourceOrder
                : int.MaxValue)) {
            var ordered = group
                .OrderBy(item => item.LogicalOrder ?? int.MaxValue)
                .ThenBy(item => item.PageNumber)
                .ThenBy(item => item.Order)
                .ThenBy(item => item.Y)
                .ThenBy(item => item.X)
                .ToList();
            var first = ordered[0];
            string renderedText = string.Concat(ordered.Where(item => !item.IsAnchor).Select(item => item.Text)).Trim();
            string anchorText = ordered.Where(item => item.IsAnchor).Select(item => item.Text.Trim()).FirstOrDefault(text => text.Length > 0) ?? string.Empty;
            string headingText = anchorText.Length > 0 ? anchorText : renderedText;
            if (headingText.Length == 0) continue;
            HtmlRenderBookmarkDefinition? definition = null;
            bookmarks?.TryGetValue(first.NodeId, out definition);
            string label = string.IsNullOrWhiteSpace(definition?.Label) ? headingText : definition!.Label!;
            headings.Add(new HtmlRenderHeading(first.NodeId, first.Level, label, first.PageNumber, first.X, first.Y, definition?.State ?? HtmlRenderBookmarkState.Default));
        }

        return headings;
    }

    private static IEnumerable<HtmlRenderTextFragment> EnumerateTextFragments(IEnumerable<HtmlRenderVisual> visuals) {
        foreach (HtmlRenderVisual visual in visuals.OrderBy(item => item.PaintOrder)) {
            if (visual is HtmlRenderSemanticGroup { Role: HtmlRenderSemanticGroupRole.Artifact }) {
                continue;
            }
            if (visual is HtmlRenderLogicalTextGroup logicalTextGroup) {
                HtmlRenderText? representative = EnumerateVisuals(logicalTextGroup.Visuals).OfType<HtmlRenderText>().FirstOrDefault();
                if (representative != null) {
                    yield return new HtmlRenderTextFragment(
                        logicalTextGroup.Text,
                        representative.SemanticRole,
                        representative.SemanticNodeId,
                        logicalTextGroup.X,
                        logicalTextGroup.Y,
                        logicalTextGroup.PaintOrder,
                        representative.SemanticFragmentOrder);
                }
                continue;
            }
            if (visual is HtmlRenderText text) {
                yield return new HtmlRenderTextFragment(text.Text, text.SemanticRole, text.SemanticNodeId, text.X, text.Y, text.PaintOrder, text.SemanticFragmentOrder);
                continue;
            }

            IEnumerable<HtmlRenderVisual>? children = ChildVisuals(visual);
            if (children == null) continue;
            foreach (HtmlRenderTextFragment fragment in EnumerateTextFragments(children)) yield return fragment;
        }
    }

    private readonly struct HtmlRenderTextFragment {
        internal HtmlRenderTextFragment(string text, string? semanticRole, int? semanticNodeId, double x, double y, int paintOrder, int? logicalOrder) {
            Text = text;
            SemanticRole = semanticRole;
            SemanticNodeId = semanticNodeId;
            X = x;
            Y = y;
            PaintOrder = paintOrder;
            LogicalOrder = logicalOrder;
        }

        internal string Text { get; }
        internal string? SemanticRole { get; }
        internal int? SemanticNodeId { get; }
        internal double X { get; }
        internal double Y { get; }
        internal int PaintOrder { get; }
        internal int? LogicalOrder { get; }
    }
}
