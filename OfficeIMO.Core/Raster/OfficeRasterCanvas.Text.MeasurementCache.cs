using System;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    private readonly struct TextMeasurementKey : IEquatable<TextMeasurementKey> {
        internal TextMeasurementKey(string text, double fontSize, string? fontFamily, OfficeFontStyle style) {
            Text = text;
            FontSize = fontSize;
            FontFamily = fontFamily ?? string.Empty;
            Style = OfficeFontFace.NormalizeStyle(style);
        }

        private string Text { get; }
        private double FontSize { get; }
        private string FontFamily { get; }
        private OfficeFontStyle Style { get; }

        public bool Equals(TextMeasurementKey other) =>
            FontSize.Equals(other.FontSize) &&
            string.Equals(Text, other.Text, StringComparison.Ordinal) &&
            string.Equals(FontFamily, other.FontFamily, StringComparison.Ordinal) &&
            Style == other.Style;

        public override bool Equals(object? obj) =>
            obj is TextMeasurementKey other && Equals(other);

        public override int GetHashCode() {
            unchecked {
                int hash = StringComparer.Ordinal.GetHashCode(Text);
                hash = (hash * 397) ^ FontSize.GetHashCode();
                hash = (hash * 397) ^ StringComparer.Ordinal.GetHashCode(FontFamily);
                hash = (hash * 397) ^ Style.GetHashCode();
                return hash;
            }
        }
    }
}
