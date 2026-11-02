namespace OfficeIMO.Drawing;

public sealed partial class OfficeDrawing {
    internal IOfficeTextShapingProvider? TextShapingProvider { get; set; }
    internal string? TextShapingLanguage { get; set; }

    // Most drawings never carry a scoped face, so the collection is created on first use.
    private OfficeFontFaceCollection? _fonts;

    /// <summary>Scoped TrueType faces embedded or rasterized with this drawing.</summary>
    public OfficeFontFaceCollection Fonts {
        get => _fonts ?? System.Threading.Interlocked.CompareExchange(ref _fonts, new OfficeFontFaceCollection(), null) ?? _fonts!;
        private set => _fonts = value;
    }

    /// <summary>Copies the scoped faces without creating a collection for a drawing that has none.</summary>
    private OfficeFontFaceCollection? CloneFonts() => _fonts?.Clone();

    /// <summary>Adds or replaces a scoped TrueType face and returns this drawing.</summary>
    public OfficeDrawing AddFont(string familyName, byte[] data, OfficeFontStyle style = OfficeFontStyle.Regular) {
        Fonts.Add(familyName, data, style);
        return this;
    }
}
