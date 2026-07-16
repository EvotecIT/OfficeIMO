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
            string? hyperlinkAnchor = null,
            bool noProof = false,
            LegacyDocFieldKind fieldKind = LegacyDocFieldKind.None,
            string? fieldInstruction = null,
            LegacyDocCharacterFormatProperties specified = LegacyDocCharacterFormatProperties.None,
            string? language = null,
            string? eastAsiaLanguage = null,
            LegacyDocPicture? picture = null,
            LegacyDocRevision revision = default)
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
                noProof,
                caps,
                verticalPosition,
                underline,
                highlight,
                fontSizeHalfPoints,
                colorHex,
                fontFamily,
                Array.Empty<int>(),
                hyperlinkUri,
                hyperlinkAnchor,
                fieldKind,
                fieldInstruction,
                specified,
                characterSpacingTwips: null,
                language: language,
                eastAsiaLanguage: eastAsiaLanguage,
                picture: picture,
                revision: revision) {
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
            bool noProof,
            LegacyDocCapsKind? caps,
            LegacyDocVerticalPositionKind? verticalPosition,
            LegacyDocUnderlineKind? underline,
            LegacyDocHighlightColorKind? highlight,
            int? fontSizeHalfPoints,
            string? colorHex,
            string? fontFamily,
            IReadOnlyList<int> characterPositions,
            string? hyperlinkUri = null,
            string? hyperlinkAnchor = null,
            LegacyDocFieldKind fieldKind = LegacyDocFieldKind.None,
            string? fieldInstruction = null,
            LegacyDocCharacterFormatProperties specified = LegacyDocCharacterFormatProperties.None,
            int? characterSpacingTwips = null,
            string? language = null,
            string? eastAsiaLanguage = null,
            LegacyDocPicture? picture = null,
            LegacyDocRevision revision = default) {
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
            NoProof = noProof;
            Caps = caps;
            VerticalPosition = verticalPosition;
            Underline = underline;
            Highlight = highlight;
            FontSizeHalfPoints = fontSizeHalfPoints;
            ColorHex = string.IsNullOrWhiteSpace(colorHex)
                ? null
                : colorHex!.Replace("#", string.Empty).ToUpperInvariant();
            FontFamily = fontFamily;
            CharacterSpacingTwips = characterSpacingTwips;
            Language = language;
            EastAsiaLanguage = eastAsiaLanguage;
            Picture = picture;
            CharacterPositions = characterPositions.Count == 0
                ? Array.Empty<int>()
                : characterPositions.ToArray();
            HyperlinkUri = string.IsNullOrWhiteSpace(hyperlinkUri) ? null : hyperlinkUri;
            HyperlinkAnchor = string.IsNullOrWhiteSpace(hyperlinkAnchor) ? null : hyperlinkAnchor;
            FieldKind = fieldKind;
            FieldInstruction = string.IsNullOrWhiteSpace(fieldInstruction) ? null : fieldInstruction;
            Specified = specified;
            Revision = revision;
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

        internal bool NoProof { get; }

        internal LegacyDocCapsKind? Caps { get; }

        internal LegacyDocVerticalPositionKind? VerticalPosition { get; }

        internal LegacyDocUnderlineKind? Underline { get; }

        internal LegacyDocHighlightColorKind? Highlight { get; }

        internal int? FontSizeHalfPoints { get; }

        internal string? ColorHex { get; }

        internal string? FontFamily { get; }

        internal int? CharacterSpacingTwips { get; }

        internal string? Language { get; }

        internal string? EastAsiaLanguage { get; }

        internal LegacyDocPicture? Picture { get; }

        internal IReadOnlyList<int> CharacterPositions { get; }

        internal string? HyperlinkUri { get; }

        internal string? HyperlinkAnchor { get; }

        internal LegacyDocFieldKind FieldKind { get; }

        internal string? FieldInstruction { get; }

        internal LegacyDocCharacterFormatProperties Specified { get; }

        internal LegacyDocRevision Revision { get; }

        internal bool IsPageNumber => FieldKind == LegacyDocFieldKind.Page;

        internal bool IsNumPages => FieldKind == LegacyDocFieldKind.NumPages;

        internal bool IsStaticDateTimeField =>
            FieldKind == LegacyDocFieldKind.Date
            || FieldKind == LegacyDocFieldKind.Time
            || FieldKind == LegacyDocFieldKind.CreateDate
            || FieldKind == LegacyDocFieldKind.SaveDate
            || FieldKind == LegacyDocFieldKind.PrintDate;

        internal bool IsStaticDisplayField =>
            IsStaticDateTimeField
            || FieldKind == LegacyDocFieldKind.DocumentProperty
            || FieldKind == LegacyDocFieldKind.Equation;

        internal bool IsSpecified(LegacyDocCharacterFormatProperties property) {
            return (Specified & property) != 0;
        }

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
