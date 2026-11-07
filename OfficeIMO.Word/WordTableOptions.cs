namespace OfficeIMO.Word;

/// <summary>Common table options with the Word table style and layout mode.</summary>
public sealed class WordTableOptions : OfficeTableOptions {
    /// <summary>Gets or sets the table style. Defaults to <see cref="WordTableStyle.TableGrid"/>.</summary>
    public WordTableStyle Style { get; set; } = WordTableStyle.TableGrid;

    /// <summary>Gets or sets the table layout mode, or <see langword="null"/> to keep the style's default layout.</summary>
    public WordTableLayoutMode? LayoutMode { get; set; }
}
