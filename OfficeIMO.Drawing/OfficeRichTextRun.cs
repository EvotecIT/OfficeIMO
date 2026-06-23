namespace OfficeIMO.Drawing;

/// <summary>
/// Describes one styled text run for shared rich text layout and rendering.
/// </summary>
public sealed class OfficeRichTextRun {
    /// <summary>
    /// Creates a styled text run.
    /// </summary>
    /// <param name="text">Run text.</param>
    /// <param name="fontSize">Font size used for measurement and rendering.</param>
    /// <param name="color">Text color.</param>
    /// <param name="bold">Whether the run should render as bold.</param>
    /// <param name="italic">Whether the run should render as italic.</param>
    /// <param name="underline">Whether the run should render with underline.</param>
    /// <param name="fontFamily">Preferred font family for SVG or future font-aware renderers.</param>
    public OfficeRichTextRun(string? text, double fontSize, OfficeColor color, bool bold = false, bool italic = false, bool underline = false, string? fontFamily = null) {
        Text = text ?? string.Empty;
        FontSize = fontSize;
        Color = color;
        Bold = bold;
        Italic = italic;
        Underline = underline;
        FontFamily = string.IsNullOrWhiteSpace(fontFamily) ? "Arial, sans-serif" : fontFamily!;
    }

    /// <summary>
    /// Gets the run text.
    /// </summary>
    public string Text { get; }

    /// <summary>
    /// Gets the run font size.
    /// </summary>
    public double FontSize { get; }

    /// <summary>
    /// Gets the run text color.
    /// </summary>
    public OfficeColor Color { get; }

    /// <summary>
    /// Gets whether the run should render as bold.
    /// </summary>
    public bool Bold { get; }

    /// <summary>
    /// Gets whether the run should render as italic.
    /// </summary>
    public bool Italic { get; }

    /// <summary>
    /// Gets whether the run should render with underline.
    /// </summary>
    public bool Underline { get; }

    /// <summary>
    /// Gets the preferred font family for SVG or future font-aware renderers.
    /// </summary>
    public string FontFamily { get; }
}
