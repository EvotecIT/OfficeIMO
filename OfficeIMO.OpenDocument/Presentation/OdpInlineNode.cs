namespace OfficeIMO.OpenDocument;

/// <summary>Kind of node in an ODP paragraph's ordered inline syntax.</summary>
public enum OdpInlineNodeKind {
    /// <summary>Plain text, including ODF space, tab, and line-break elements.</summary>
    Text,
    /// <summary>A styled text run.</summary>
    Run,
    /// <summary>A hyperlink.</summary>
    Hyperlink,
    /// <summary>An inline element not represented by the current typed surface.</summary>
    Other
}

/// <summary>
/// An ordered typed direct child in an ODP paragraph. Nested inline markup is surfaced
/// as <see cref="OdpInlineNodeKind.Other"/> so conversion loss remains explicit.
/// </summary>
public sealed class OdpInlineNode {
    private OdpInlineNode(OdpInlineNodeKind kind, string text, OdpRun? run = null,
        OdpHyperlink? hyperlink = null, string? qualifiedName = null) {
        Kind = kind;
        Text = text;
        Run = run;
        Hyperlink = hyperlink;
        QualifiedName = qualifiedName;
    }

    /// <summary>Node kind.</summary>
    public OdpInlineNodeKind Kind { get; }
    /// <summary>Decoded text contributed by this node.</summary>
    public string Text { get; }
    /// <summary>Styled run for <see cref="OdpInlineNodeKind.Run"/>.</summary>
    public OdpRun? Run { get; }
    /// <summary>Hyperlink for <see cref="OdpInlineNodeKind.Hyperlink"/>.</summary>
    public OdpHyperlink? Hyperlink { get; }
    /// <summary>Expanded XML name for an unrepresented element.</summary>
    public string? QualifiedName { get; }

    internal static IReadOnlyList<OdpInlineNode> Read(OdpPresentation presentation, XElement paragraph) {
        _ = OdfTextCodec.Read(paragraph);
        var result = new List<OdpInlineNode>();
        var plainNodes = new List<XNode>();

        void FlushPlain() {
            if (plainNodes.Count == 0) return;
            string text = OdfTextCodec.ReadNodes(plainNodes);
            if (text.Length > 0) result.Add(new OdpInlineNode(OdpInlineNodeKind.Text, text));
            plainNodes.Clear();
        }

        foreach (XNode node in paragraph.Nodes()) {
            if (node is XText) { plainNodes.Add(node); continue; }
            if (!(node is XElement element)) continue;
            if (element.Name == OdfNamespaces.Text + "s"
                || element.Name == OdfNamespaces.Text + "tab"
                || element.Name == OdfNamespaces.Text + "line-break") {
                plainNodes.Add(element);
                continue;
            }
            FlushPlain();
            if ((element.Name == OdfNamespaces.Text + "span"
                    || element.Name == OdfNamespaces.Text + "a")
                && HasNestedInlineMarkup(element)) {
                result.Add(new OdpInlineNode(OdpInlineNodeKind.Other, OdfTextCodec.Read(element),
                    qualifiedName: element.Name.ToString()));
            } else if (element.Name == OdfNamespaces.Text + "span") {
                var run = new OdpRun(presentation, element);
                result.Add(new OdpInlineNode(OdpInlineNodeKind.Run, run.Text, run: run));
            } else if (element.Name == OdfNamespaces.Text + "a") {
                var hyperlink = new OdpHyperlink(presentation, element);
                result.Add(new OdpInlineNode(OdpInlineNodeKind.Hyperlink, hyperlink.Text, hyperlink: hyperlink));
            } else {
                result.Add(new OdpInlineNode(OdpInlineNodeKind.Other, OdfTextCodec.Read(element),
                    qualifiedName: element.Name.ToString()));
            }
        }
        FlushPlain();
        return result;
    }

    private static bool HasNestedInlineMarkup(XElement element) => element.Elements().Any(child =>
        child.Name != OdfNamespaces.Text + "s"
        && child.Name != OdfNamespaces.Text + "tab"
        && child.Name != OdfNamespaces.Text + "line-break");
}

/// <summary>An XML-backed ODP hyperlink. Targets are preserved and never fetched.</summary>
public sealed class OdpHyperlink {
    private readonly OdpPresentation _presentation;
    private readonly XElement _element;

    internal OdpHyperlink(OdpPresentation presentation, XElement element) {
        _presentation = presentation;
        _element = element;
    }

    /// <summary>Decoded display text.</summary>
    public string Text { get => OdfTextCodec.Read(_element); set { OdfTextCodec.Replace(_element, value); Dirty(); } }
    /// <summary>Link target.</summary>
    public string Href {
        get => (string?)_element.Attribute(OdfNamespaces.XLink + "href") ?? string.Empty;
        set {
            if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("Hyperlink target cannot be empty.", nameof(value));
            _element.SetAttributeValue(OdfNamespaces.XLink + "href", value);
            Dirty();
        }
    }
    /// <summary>ODF target frame behavior, if authored.</summary>
    public string? TargetFrameName {
        get => (string?)_element.Attribute(OdfNamespaces.Office + "target-frame-name");
        set { _element.SetAttributeValue(OdfNamespaces.Office + "target-frame-name", NormalizeOptional(value)); Dirty(); }
    }
    /// <summary>Raw XLink show behavior, if authored.</summary>
    public string? ShowBehavior {
        get => (string?)_element.Attribute(OdfNamespaces.XLink + "show");
        set { _element.SetAttributeValue(OdfNamespaces.XLink + "show", NormalizeOptional(value)); Dirty(); }
    }
    /// <summary>Referenced text style name.</summary>
    public string? StyleName { get => (string?)_element.Attribute(OdfNamespaces.Text + "style-name"); set { _element.SetAttributeValue(OdfNamespaces.Text + "style-name", value); Dirty(); } }
    /// <summary>Explicit or inherited bold state.</summary>
    public bool? Bold { get => Resolve(style => style.Bold); set => EnsureStyle().Bold = value; }
    /// <summary>Explicit or inherited italic state.</summary>
    public bool? Italic { get => Resolve(style => style.Italic); set => EnsureStyle().Italic = value; }
    /// <summary>Explicit or inherited underline state.</summary>
    public bool? Underline { get => Resolve(style => style.Underline); set => EnsureStyle().Underline = value; }
    /// <summary>Whether the effective underline uses a non-solid ODF decoration style.</summary>
    public bool UsesNonSolidUnderlineStyle => Resolve(style => style.UsesNonSolidUnderlineStyle) == true;
    /// <summary>Explicit or inherited strike-through state.</summary>
    public bool? StrikeThrough { get => Resolve(style => style.StrikeThrough); set => EnsureStyle().StrikeThrough = value; }
    /// <summary>Whether the effective line-through uses a non-solid ODF decoration style.</summary>
    public bool UsesNonSolidLineThroughStyle => Resolve(style => style.UsesNonSolidLineThroughStyle) == true;
    /// <summary>Explicit or inherited font size.</summary>
    public OdfLength? FontSize { get => Resolve(style => style.FontSize); set => EnsureStyle().FontSize = value; }
    /// <summary>Explicit or inherited font family.</summary>
    public string? FontFamily { get => ResolveReference(style => style.FontFamily); set => EnsureStyle().FontFamily = value; }
    /// <summary>Explicit or inherited text color.</summary>
    public OdfColor? Color { get => Resolve(style => style.Color); set => EnsureStyle().Color = value; }
    /// <summary>Explicit or inherited text background color.</summary>
    public OdfColor? BackgroundColor {
        get {
            OdfStyle? style = StyleName == null ? null : _presentation.Styles.Find(
                OdfStyleFamily.Text, StyleName);
            return _presentation.Styles.ResolveTextBackgroundColor(style);
        }
        set => EnsureStyle().TextBackgroundColor = value;
    }

    private OdfStyle EnsureStyle() => _presentation.Styles.EnsureAutomaticStyle(
        _element, OdfNamespaces.Text + "style-name", OdfStyleFamily.Text, "ofLink");
    private T? Resolve<T>(Func<OdfStyle, T?> selector) where T : struct {
        OdfStyle? style = StyleName == null ? null : _presentation.Styles.Find(OdfStyleFamily.Text, StyleName); if (style == null) return null;
        foreach (OdfStyle candidate in _presentation.Styles.Resolve(style)) { T? value = selector(candidate); if (value.HasValue) return value; } return null;
    }
    private string? ResolveReference(Func<OdfStyle, string?> selector) {
        OdfStyle? style = StyleName == null ? null : _presentation.Styles.Find(OdfStyleFamily.Text, StyleName); if (style == null) return null;
        foreach (OdfStyle candidate in _presentation.Styles.Resolve(style)) { string? value = selector(candidate); if (value != null) return value; } return null;
    }
    private static string? NormalizeOptional(string? value) => string.IsNullOrWhiteSpace(value) ? null : value;
    private void Dirty() => _presentation.MarkPartDirty("content.xml");
}
