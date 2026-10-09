namespace OfficeIMO.OpenDocument;

/// <summary>A native paragraph or heading in Draw text content.</summary>
public sealed partial class OdfTextParagraph : OdfTextContent {
    internal OdfTextParagraph(OdfDocument document, XElement element, XElement graphic) : base(document, element, graphic, OdfStyleFamily.Paragraph) { }
    /// <summary>Whether this element is a heading.</summary>
    public bool IsHeading => Element.Name == OdfNamespaces.Text + "h";
    /// <summary>Native horizontal alignment token, such as start, center, end, or justify.</summary>
    public string? TextAlign { get => Styles.Select(s => s.TextAlign).FirstOrDefault(v => v != null); set => EnsureStyle().TextAlign = value; }
    /// <summary>Native writing-mode token, such as lr-tb or rl-tb.</summary>
    public string? WritingMode { get => Styles.Select(s => s.WritingMode).FirstOrDefault(v => v != null); set => EnsureStyle().WritingMode = value; }
    /// <summary>Effective paragraph line height, absolute or percentage.</summary>
    public OdfLength? LineHeight { get => Resolve(s => s.LineHeight); set => EnsureStyle().LineHeight = value; }
    /// <summary>Effective left paragraph margin.</summary>
    public OdfLength? MarginLeft { get => Resolve(s => s.MarginLeft); set => EnsureStyle().MarginLeft = value; }
    /// <summary>Effective right paragraph margin.</summary>
    public OdfLength? MarginRight { get => Resolve(s => s.MarginRight); set => EnsureStyle().MarginRight = value; }
    /// <summary>Effective top paragraph margin.</summary>
    public OdfLength? MarginTop { get => Resolve(s => s.MarginTop); set => EnsureStyle().MarginTop = value; }
    /// <summary>Effective bottom paragraph margin.</summary>
    public OdfLength? MarginBottom { get => Resolve(s => s.MarginBottom); set => EnsureStyle().MarginBottom = value; }
}
