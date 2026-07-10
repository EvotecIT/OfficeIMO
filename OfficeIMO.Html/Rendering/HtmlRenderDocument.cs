using System.Collections.ObjectModel;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

/// <summary>
/// Backend-neutral result shared by HTML image and PDF export.
/// </summary>
public sealed class HtmlRenderDocument {
    private readonly ReadOnlyCollection<HtmlRenderPage> _pages;
    private readonly OfficeFontFaceCollection _fonts;

    internal HtmlRenderDocument(HtmlRenderMode mode, IEnumerable<HtmlRenderPage> pages, HtmlDiagnosticReport diagnostics, OfficeFontFaceCollection? fonts = null) {
        Mode = mode;
        _pages = new List<HtmlRenderPage>(pages ?? throw new ArgumentNullException(nameof(pages))).AsReadOnly();
        if (_pages.Count == 0) {
            throw new ArgumentException("A rendered HTML document requires at least one surface.", nameof(pages));
        }

        Diagnostics = (diagnostics ?? throw new ArgumentNullException(nameof(diagnostics))).Clone();
        _fonts = fonts?.Clone() ?? new OfficeFontFaceCollection();
    }

    /// <summary>Layout mode used to produce the result.</summary>
    public HtmlRenderMode Mode { get; }

    /// <summary>Rendered pages, or one page for continuous output.</summary>
    public IReadOnlyList<HtmlRenderPage> Pages => _pages;

    /// <summary>Diagnostics emitted while parsing, laying out, and preparing paint operations.</summary>
    public HtmlDiagnosticReport Diagnostics { get; }

    /// <summary>Independent snapshot of scoped font faces retained for image and PDF backends.</summary>
    public OfficeFontFaceCollection Fonts => _fonts.Clone();

    /// <summary>Concatenated searchable text retained by the shared render model.</summary>
    public string Text => string.Join("\n", _pages.SelectMany(page => EnumerateVisuals(page.Visuals)).OfType<HtmlRenderText>().Select(text => text.Text));

    private static IEnumerable<HtmlRenderVisual> EnumerateVisuals(IEnumerable<HtmlRenderVisual> visuals) {
        foreach (HtmlRenderVisual visual in visuals) {
            yield return visual;
            IEnumerable<HtmlRenderVisual>? children = visual is HtmlRenderClipGroup clipGroup
                ? clipGroup.Visuals
                : visual is HtmlRenderEffectGroup effectGroup ? effectGroup.Visuals : null;
            if (children == null) continue;
            foreach (HtmlRenderVisual child in EnumerateVisuals(children)) yield return child;
        }
    }
}
