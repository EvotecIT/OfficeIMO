using System;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    private readonly struct TextMeasurementKey : IEquatable<TextMeasurementKey> {
        internal TextMeasurementKey(string text, double fontSize, string? fontFamily, OfficeFontStyle style,
            bool paintedGlyphOrder = false, OfficeFontFaceDescriptor? face = null) {
            Text = text;
            PaintedGlyphOrder = paintedGlyphOrder;
            FontSize = fontSize;
            FontFamily = fontFamily ?? string.Empty;
            Style = OfficeFontFace.NormalizeStyle(style);
            Face = face ?? OfficeFontFaceDescriptor.FromStyle(style);
        }

        private string Text { get; }
        private double FontSize { get; }
        private string FontFamily { get; }
        private OfficeFontStyle Style { get; }
        private OfficeFontFaceDescriptor Face { get; }
        // Painted glyph runs skip shaping, so their measurements differ from shaped text.
        private bool PaintedGlyphOrder { get; }

        public bool Equals(TextMeasurementKey other) =>
            FontSize.Equals(other.FontSize) &&
            string.Equals(Text, other.Text, StringComparison.Ordinal) &&
            string.Equals(FontFamily, other.FontFamily, StringComparison.Ordinal) &&
            Style == other.Style && Face == other.Face &&
            PaintedGlyphOrder == other.PaintedGlyphOrder;

        public override bool Equals(object? obj) =>
            obj is TextMeasurementKey other && Equals(other);

        public override int GetHashCode() {
            unchecked {
                int hash = StringComparer.Ordinal.GetHashCode(Text);
                hash = (hash * 397) ^ FontSize.GetHashCode();
                hash = (hash * 397) ^ StringComparer.Ordinal.GetHashCode(FontFamily);
                hash = (hash * 397) ^ Style.GetHashCode();
                hash = (hash * 397) ^ Face.GetHashCode();
                hash = (hash * 397) ^ (PaintedGlyphOrder ? 1 : 0);
                return hash;
            }
        }
    }
}
