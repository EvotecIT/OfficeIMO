namespace OfficeIMO.Word.LegacyDoc.Model {
    internal sealed class LegacyDocTextRun {
        internal LegacyDocTextRun(
            string text,
            bool bold,
            bool italic,
            bool strike,
            bool doubleStrike,
            bool outline,
            bool shadow,
            bool emboss,
            bool imprint,
            bool hidden,
            LegacyDocCapsKind? caps,
            LegacyDocVerticalPositionKind? verticalPosition,
            LegacyDocUnderlineKind? underline,
            LegacyDocHighlightColorKind? highlight,
            int? fontSizeHalfPoints,
            string? colorHex,
            string? fontFamily) {
            Text = text;
            Bold = bold;
            Italic = italic;
            Strike = strike;
            DoubleStrike = doubleStrike;
            Outline = outline;
            Shadow = shadow;
            Emboss = emboss;
            Imprint = imprint;
            Hidden = hidden;
            Caps = caps;
            VerticalPosition = verticalPosition;
            Underline = underline;
            Highlight = highlight;
            FontSizeHalfPoints = fontSizeHalfPoints;
            ColorHex = colorHex;
            FontFamily = fontFamily;
        }

        internal string Text { get; }

        internal bool Bold { get; }

        internal bool Italic { get; }

        internal bool Strike { get; }

        internal bool DoubleStrike { get; }

        internal bool Outline { get; }

        internal bool Shadow { get; }

        internal bool Emboss { get; }

        internal bool Imprint { get; }

        internal bool Hidden { get; }

        internal LegacyDocCapsKind? Caps { get; }

        internal LegacyDocVerticalPositionKind? VerticalPosition { get; }

        internal LegacyDocUnderlineKind? Underline { get; }

        internal LegacyDocHighlightColorKind? Highlight { get; }

        internal int? FontSizeHalfPoints { get; }

        internal string? ColorHex { get; }

        internal string? FontFamily { get; }
    }
}
