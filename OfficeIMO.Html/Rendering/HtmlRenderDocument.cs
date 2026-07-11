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

    internal HtmlRenderDocument(HtmlRenderMode mode, IEnumerable<HtmlRenderPage> pages, HtmlDiagnosticReport diagnostics, OfficeFontFaceCollection? fonts = null, HtmlRenderMetadata? metadata = null) {
        Mode = mode;
        _pages = new List<HtmlRenderPage>(pages ?? throw new ArgumentNullException(nameof(pages))).AsReadOnly();
        if (_pages.Count == 0) {
            throw new ArgumentException("A rendered HTML document requires at least one surface.", nameof(pages));
        }

        Diagnostics = (diagnostics ?? throw new ArgumentNullException(nameof(diagnostics))).Clone();
        _fonts = fonts?.Clone() ?? new OfficeFontFaceCollection();
        Metadata = metadata ?? new HtmlRenderMetadata(null, null);
        _headings = BuildHeadings(_pages).AsReadOnly();
    }

    /// <summary>Layout mode used to produce the result.</summary>
    public HtmlRenderMode Mode { get; }

    /// <summary>Rendered pages, or one page for continuous output.</summary>
    public IReadOnlyList<HtmlRenderPage> Pages => _pages;

    /// <summary>Diagnostics emitted while parsing, laying out, and preparing paint operations.</summary>
    public HtmlDiagnosticReport Diagnostics { get; }

    /// <summary>Independent snapshot of scoped font faces retained for image and PDF backends.</summary>
    public OfficeFontFaceCollection Fonts => _fonts.Clone();

    /// <summary>Source document metadata retained for image and PDF adapters.</summary>
    public HtmlRenderMetadata Metadata { get; }

    /// <summary>Source headings retained in document order for navigation-capable backends.</summary>
    public IReadOnlyList<HtmlRenderHeading> Headings => _headings;

    /// <summary>Concatenated searchable text retained by the shared render model.</summary>
    public string Text => string.Join("\n", _pages.SelectMany(page => EnumerateVisuals(page.Scene)).OfType<HtmlRenderText>().Select(text => text.Text));

    private static IEnumerable<HtmlRenderVisual> EnumerateVisuals(IEnumerable<HtmlRenderVisual> visuals) {
        foreach (HtmlRenderVisual visual in visuals) {
            yield return visual;
            IEnumerable<HtmlRenderVisual>? children = visual is HtmlRenderClipGroup clipGroup
                ? clipGroup.Visuals
                : visual is HtmlRenderPathClipGroup pathClipGroup
                    ? pathClipGroup.Visuals
                : visual is HtmlRenderEffectGroup effectGroup ? effectGroup.Visuals
                : visual is HtmlRenderSemanticGroup semanticGroup ? semanticGroup.Visuals : null;
            if (children == null) continue;
            foreach (HtmlRenderVisual child in EnumerateVisuals(children)) yield return child;
        }
    }

    private static List<HtmlRenderHeading> BuildHeadings(IReadOnlyList<HtmlRenderPage> pages) {
        var fragments = new List<(int NodeId, int Level, string Text, int PageNumber, double X, double Y, int Order)>();
        foreach (HtmlRenderPage page in pages) {
            foreach (HtmlRenderText text in EnumerateVisuals(page.Scene).OfType<HtmlRenderText>()) {
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
}
