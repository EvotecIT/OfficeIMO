namespace OfficeIMO.Word.LegacyDoc.Model {
    internal sealed class LegacyDocTextRun {
        internal LegacyDocTextRun(
            string text,
            bool bold,
            bool italic,
            LegacyDocUnderlineKind? underline,
            int? fontSizeHalfPoints,
            string? colorHex) {
            Text = text;
            Bold = bold;
            Italic = italic;
            Underline = underline;
            FontSizeHalfPoints = fontSizeHalfPoints;
            ColorHex = colorHex;
        }

        internal string Text { get; }

        internal bool Bold { get; }

        internal bool Italic { get; }

        internal LegacyDocUnderlineKind? Underline { get; }

        internal int? FontSizeHalfPoints { get; }

        internal string? ColorHex { get; }
    }
}
