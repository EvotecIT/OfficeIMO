namespace OfficeIMO.Word.LegacyDoc.Model {
    internal sealed class LegacyDocTextRun {
        internal LegacyDocTextRun(
            string text,
            bool bold,
            bool italic,
            bool strike,
            LegacyDocVerticalPositionKind? verticalPosition,
            LegacyDocUnderlineKind? underline,
            int? fontSizeHalfPoints,
            string? colorHex,
            string? fontFamily) {
            Text = text;
            Bold = bold;
            Italic = italic;
            Strike = strike;
            VerticalPosition = verticalPosition;
            Underline = underline;
            FontSizeHalfPoints = fontSizeHalfPoints;
            ColorHex = colorHex;
            FontFamily = fontFamily;
        }

        internal string Text { get; }

        internal bool Bold { get; }

        internal bool Italic { get; }

        internal bool Strike { get; }

        internal LegacyDocVerticalPositionKind? VerticalPosition { get; }

        internal LegacyDocUnderlineKind? Underline { get; }

        internal int? FontSizeHalfPoints { get; }

        internal string? ColorHex { get; }

        internal string? FontFamily { get; }
    }
}
