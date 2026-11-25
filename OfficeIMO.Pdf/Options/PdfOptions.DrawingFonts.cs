using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfCreateOptions {
    private OfficeFontFaceCollection? _drawingProfileFonts;

    /// <summary>Gets an independent scope for drawing layout using fonts explicitly supplied by the caller.</summary>
    internal OfficeFontFaceCollection? CloneDrawingProfileFonts() => _drawingProfileFonts?.Clone();

    internal OfficeDrawingText ResolveDrawingFont(OfficeDrawingText text) {
        if (_drawingProfileFonts == null || !_drawingProfileFonts.TryResolveFaceForText(
            text.Text, text.Font.FamilyName, text.Font.Face, out OfficeFontFace? face)
            || face == null || _renderingProfileOwnedNamedFamilyNames?.Contains(face.ResourceFamilyName) != true) return text;
        // The unique resource family carries the chosen program through PDF's legacy style slots.
        OfficeFontStyle decorations = text.Font.Style & ~(OfficeFontStyle.Bold | OfficeFontStyle.Italic);
        return text.CloneWithFont(new OfficeFontInfo(face.ResourceFamilyName, text.Font.Size, face.Descriptor, decorations));
    }
}
