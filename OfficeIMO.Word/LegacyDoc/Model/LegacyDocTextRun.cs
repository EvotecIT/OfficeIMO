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
            string? fontFamily,
            string? hyperlinkUri = null,
            string? hyperlinkAnchor = null)
            : this(
                text,
                bold,
                italic,
                strike,
                doubleStrike,
                outline,
                shadow,
                emboss,
                imprint,
                hidden,
                caps,
                verticalPosition,
                underline,
                highlight,
                fontSizeHalfPoints,
                colorHex,
                fontFamily,
                Array.Empty<int>(),
                hyperlinkUri,
                hyperlinkAnchor) {
        }

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
            string? fontFamily,
            IReadOnlyList<int> characterPositions,
            string? hyperlinkUri = null,
            string? hyperlinkAnchor = null) {
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
            CharacterPositions = characterPositions.Count == 0
                ? Array.Empty<int>()
                : characterPositions.ToArray();
            HyperlinkUri = string.IsNullOrWhiteSpace(hyperlinkUri) ? null : hyperlinkUri;
            HyperlinkAnchor = string.IsNullOrWhiteSpace(hyperlinkAnchor) ? null : hyperlinkAnchor;
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

        internal IReadOnlyList<int> CharacterPositions { get; }

        internal string? HyperlinkUri { get; }

        internal string? HyperlinkAnchor { get; }

        internal LegacyDocHyperlinkTarget HyperlinkTarget {
            get {
                if (HyperlinkUri != null) {
                    return LegacyDocHyperlinkTarget.ForUri(HyperlinkUri);
                }

                if (HyperlinkAnchor != null) {
                    return LegacyDocHyperlinkTarget.ForAnchor(HyperlinkAnchor);
                }

                return default;
            }
        }
    }
}
