namespace OfficeIMO.IWork;

/// <summary>An immutable character run for template-free native iWork creation.</summary>
public sealed class IWorkTextRunData {
    /// <summary>Creates a run with one font, size, opaque sRGB color and character decoration.</summary>
    /// <remarks>CR and CRLF normalize to LF. Native font availability and wrapping depend on the application.</remarks>
    public IWorkTextRunData(string text, string fontName = "Arial", float fontSizePoints = 32,
        string color = "000000", bool bold = false, bool italic = false, bool underline = false, bool strikethrough = false)
        : this(text, fontName, fontSizePoints, IWorkCreationValidation.ParseColor(color), bold, italic, underline, strikethrough) { }

    internal IWorkTextRunData(string text, string fontName, float fontSizePoints, IWorkColor color,
        bool bold = false, bool italic = false, bool underline = false, bool strikethrough = false, bool allowTabs = false) {
        Text = IWorkCreationValidation.NormalizeText(text, allowTabs);
        IWorkCreationValidation.ValidateFontName(fontName);
        IWorkCreationValidation.ValidatePositive(fontSizePoints, nameof(fontSizePoints));
        FontName = fontName; FontSizePoints = fontSizePoints; Color = color;
        Bold = bold; Italic = italic; Underline = underline; Strikethrough = strikethrough;
    }

    /// <summary>Gets text with normalized LF line breaks.</summary>
    public string Text { get; }
    /// <summary>Gets the requested native font name.</summary>
    public string FontName { get; }
    /// <summary>Gets the font size in points.</summary>
    public float FontSizePoints { get; }
    /// <summary>Gets the opaque sRGB foreground color.</summary>
    public IWorkColor Color { get; }
    /// <summary>Gets whether the run is bold.</summary>
    public bool Bold { get; }
    /// <summary>Gets whether the run is italic.</summary>
    public bool Italic { get; }
    /// <summary>Gets whether the run has a single underline.</summary>
    public bool Underline { get; }
    /// <summary>Gets whether the run has a single strikethrough.</summary>
    public bool Strikethrough { get; }

    internal bool HasDecoration => Bold || Italic || Underline || Strikethrough;
}
