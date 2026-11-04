namespace OfficeIMO.Xps;

/// <summary>A live handle to a native visual created on an XPS page, including image-brush paths.</summary>
public sealed class XpsVisual {
    internal XElement Element;
    internal XpsVisual(XpsPage page, XElement element) { Page = page; Element = element; }
    /// <summary>Gets the page containing this visual.</summary>
    public XpsPage Page { get; }
    /// <summary>Returns a detached copy of the visual's current native XML.</summary>
    public XElement GetMarkup() { Page.CheckVisual(this); return new XElement(Element); }
    /// <summary>Replaces this visual in the same native element kind, validating the complete page before committing.</summary>
    /// <remarks>Page-wide markup replacement invalidates previously returned handles.</remarks>
    public void ReplaceMarkup(XElement markup) => Page.ReplaceVisual(this, markup);
}

public sealed partial class XpsPage {
    private readonly List<XpsVisual> _visuals = new();
    private XpsVisual TrackVisual(XElement element) {
        var visual = new XpsVisual(this, element); _visuals.Add(visual); return visual;
    }
    internal void CheckVisual(XpsVisual visual) {
        if (visual.Page != this || visual.Element.Parent != _markup)
            throw new InvalidOperationException("The visual is no longer attached to this page.");
    }
    internal void ReplaceVisual(XpsVisual visual, XElement markup) {
        ArgumentNullException.ThrowIfNull(markup); CheckVisual(visual);
        if (markup.Name != visual.Element.Name) throw new ArgumentException("Expected the same native visual element and dialect.", nameof(markup));
        XElement[] originals = _markup.Elements().ToArray();
        var live = _visuals.Where(v => v.Element.Parent == _markup)
            .Select(v => (Visual: v, Index: Array.IndexOf(originals, v.Element))).ToArray();
        var candidate = GetMarkup();
        candidate.Elements().ElementAt(Array.IndexOf(originals, visual.Element)).ReplaceWith(new XElement(markup));
        Document.CommitPageContent(this, Document.ValidatePageMarkup(candidate));
        XElement[] replacements = _markup.Elements().ToArray();
        foreach (var item in live) {
            item.Visual.Element = replacements[item.Index];
            _visuals.Add(item.Visual);
        }
    }
}
