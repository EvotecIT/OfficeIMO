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

    internal HtmlRenderDocument(HtmlRenderMode mode, IEnumerable<HtmlRenderPage> pages, HtmlDiagnosticReport diagnostics, OfficeFontFaceCollection? fonts = null, HtmlRenderMetadata? metadata = null) {
        Mode = mode;
        _pages = new List<HtmlRenderPage>(pages ?? throw new ArgumentNullException(nameof(pages))).AsReadOnly();
        if (_pages.Count == 0) {
            throw new ArgumentException("A rendered HTML document requires at least one surface.", nameof(pages));
        }

        _diagnosticReport = (diagnostics ?? throw new ArgumentNullException(nameof(diagnostics))).Clone();
        _fonts = fonts?.Clone() ?? new OfficeFontFaceCollection();
        Metadata = metadata ?? new HtmlRenderMetadata(null, null);
        _headings = BuildHeadings(_pages).AsReadOnly();
    }

    /// <summary>Layout mode used to produce the result.</summary>
    public HtmlRenderMode Mode { get; }

    /// <summary>Rendered pages, or one page for continuous output.</summary>
    public IReadOnlyList<HtmlRenderPage> Pages => _pages;

    /// <summary>Diagnostics emitted while parsing, laying out, and preparing paint operations.</summary>
    public IReadOnlyList<HtmlDiagnostic> Diagnostics => _diagnosticReport.Diagnostics;

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
        foreach (HtmlRenderVisual visual in visuals.OrderBy(item => item.PaintOrder)) {
            if (visual is HtmlRenderLogicalTextGroup logicalTextGroup) {
                yield return logicalTextGroup.Text;
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

    private static IEnumerable<HtmlRenderVisual> EnumerateVisuals(IEnumerable<HtmlRenderVisual> visuals) {
        foreach (HtmlRenderVisual visual in visuals) {
            yield return visual;
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
        : visual is HtmlRenderLogicalTextGroup logicalTextGroup ? logicalTextGroup.Visuals : null;

    private static List<HtmlRenderHeading> BuildHeadings(IReadOnlyList<HtmlRenderPage> pages) {
        var fragments = new List<(int NodeId, int Level, string Text, int PageNumber, double X, double Y, int Order)>();
        foreach (HtmlRenderPage page in pages) {
            foreach (HtmlRenderTextFragment text in EnumerateTextFragments(page.Scene)) {
                if (!text.SemanticNodeId.HasValue || !HtmlRenderHeading.TryGetLevel(text.SemanticRole, out int level)) continue;
                fragments.Add((text.SemanticNodeId.Value, level, text.Text, page.PageNumber, text.X, text.Y, text.PaintOrder));
            }
        }

        var headings = new List<HtmlRenderHeading>();
        foreach (IGrouping<int, (int NodeId, int Level, string Text, int PageNumber, double X, double Y, int Order)> group in fragments
            .OrderBy(item => item.PageNumber)
            .ThenBy(item => item.Order)
            .GroupBy(item => item.NodeId)) {
            var ordered = group.OrderBy(item => item.PageNumber).ThenBy(item => item.Order).ToList();
            var first = ordered[0];
            string headingText = string.Concat(ordered.Select(item => item.Text)).Trim();
            if (headingText.Length == 0) continue;
            headings.Add(new HtmlRenderHeading(first.NodeId, first.Level, headingText, first.PageNumber, first.X, first.Y));
        }

        return headings;
    }

    private static IEnumerable<HtmlRenderTextFragment> EnumerateTextFragments(IEnumerable<HtmlRenderVisual> visuals) {
        foreach (HtmlRenderVisual visual in visuals.OrderBy(item => item.PaintOrder)) {
            if (visual is HtmlRenderLogicalTextGroup logicalTextGroup) {
                HtmlRenderText? representative = EnumerateVisuals(logicalTextGroup.Visuals).OfType<HtmlRenderText>().FirstOrDefault();
                if (representative != null) {
                    yield return new HtmlRenderTextFragment(
                        logicalTextGroup.Text,
                        representative.SemanticRole,
                        representative.SemanticNodeId,
                        logicalTextGroup.X,
                        logicalTextGroup.Y,
                        logicalTextGroup.PaintOrder);
                }
                continue;
            }
            if (visual is HtmlRenderText text) {
                yield return new HtmlRenderTextFragment(text.Text, text.SemanticRole, text.SemanticNodeId, text.X, text.Y, text.PaintOrder);
                continue;
            }

            IEnumerable<HtmlRenderVisual>? children = ChildVisuals(visual);
            if (children == null) continue;
            foreach (HtmlRenderTextFragment fragment in EnumerateTextFragments(children)) yield return fragment;
        }
    }

    private readonly struct HtmlRenderTextFragment {
        internal HtmlRenderTextFragment(string text, string? semanticRole, int? semanticNodeId, double x, double y, int paintOrder) {
            Text = text;
            SemanticRole = semanticRole;
            SemanticNodeId = semanticNodeId;
            X = x;
            Y = y;
            PaintOrder = paintOrder;
        }

        internal string Text { get; }
        internal string? SemanticRole { get; }
        internal int? SemanticNodeId { get; }
        internal double X { get; }
        internal double Y { get; }
        internal int PaintOrder { get; }
    }
}
