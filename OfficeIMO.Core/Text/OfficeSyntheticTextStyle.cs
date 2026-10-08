namespace OfficeIMO.Drawing;

// Positioned outline paint and its measurement share these drawing-unit transforms.
internal static class OfficeSyntheticTextStyle {
    internal static double BoldOffset(double size) => size / 24D;
    internal static double ItalicOffset(double baseline, double y) => (baseline - y) * 0.18D;
}
