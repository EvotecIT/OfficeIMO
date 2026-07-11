namespace OfficeIMO.Drawing;

public sealed partial class OfficeDrawing {
    /// <summary>Scoped TrueType faces embedded or rasterized with this drawing.</summary>
    public OfficeFontFaceCollection Fonts { get; } = new OfficeFontFaceCollection();

    /// <summary>Adds or replaces a scoped TrueType face and returns this drawing.</summary>
    public OfficeDrawing AddFont(string familyName, byte[] data, OfficeFontStyle style = OfficeFontStyle.Regular) {
        Fonts.Add(familyName, data, style);
        return this;
    }
}
