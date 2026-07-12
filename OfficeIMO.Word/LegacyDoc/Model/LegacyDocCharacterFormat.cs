namespace OfficeIMO.Word.LegacyDoc.Model {
    [Flags]
    internal enum LegacyDocCharacterFormatProperties {
        None = 0,
        Bold = 1 << 0,
        Italic = 1 << 1,
        Strike = 1 << 2,
        DoubleStrike = 1 << 3,
        Outline = 1 << 4,
        Shadow = 1 << 5,
        Emboss = 1 << 6,
        Imprint = 1 << 7,
        Hidden = 1 << 8,
        NoProof = 1 << 9,
        Caps = 1 << 10,
        SmallCaps = 1 << 11,
        VerticalPosition = 1 << 12,
        Underline = 1 << 13,
        Highlight = 1 << 14,
        FontSize = 1 << 15,
        Color = 1 << 16,
        FontFamily = 1 << 17,
        CharacterSpacing = 1 << 18,
        Language = 1 << 19
    }

    internal readonly struct LegacyDocCharacterFormat : IEquatable<LegacyDocCharacterFormat> {
        internal LegacyDocCharacterFormat(
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
            int? characterSpacingTwips,
            string? language,
            string? eastAsiaLanguage,
            LegacyDocCharacterFormatProperties specified = LegacyDocCharacterFormatProperties.None) {
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
            Specified = specified;
        }

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

        internal LegacyDocCharacterFormatProperties Specified { get; }

        internal bool HasFormatting =>
            Bold
            || Italic
            || Strike
            || DoubleStrike
            || Outline
            || Shadow
            || Emboss
            || Imprint
            || Hidden
            || NoProof
            || Caps != null
            || VerticalPosition != null
            || Underline != null
            || Highlight != null
            || FontSizeHalfPoints != null
            || ColorHex != null
            || FontFamily != null
            || CharacterSpacingTwips != null
            || Language != null
            || EastAsiaLanguage != null
            || Specified != LegacyDocCharacterFormatProperties.None;

        internal static LegacyDocCharacterFormat Default { get; } = new LegacyDocCharacterFormat(false, false, false, false, false, false, false, false, false, false, null, null, null, null, null, null, null, null, null, null);

        internal bool IsSpecified(LegacyDocCharacterFormatProperties property) {
            return (Specified & property) != 0;
        }

        public bool Equals(LegacyDocCharacterFormat other) {
            return Bold == other.Bold
                && Italic == other.Italic
                && Strike == other.Strike
                && DoubleStrike == other.DoubleStrike
                && Outline == other.Outline
                && Shadow == other.Shadow
                && Emboss == other.Emboss
                && Imprint == other.Imprint
                && Hidden == other.Hidden
                && NoProof == other.NoProof
                && Caps == other.Caps
                && VerticalPosition == other.VerticalPosition
                && Underline == other.Underline
                && Highlight == other.Highlight
                && FontSizeHalfPoints == other.FontSizeHalfPoints
                && string.Equals(ColorHex, other.ColorHex, StringComparison.OrdinalIgnoreCase)
                && string.Equals(FontFamily, other.FontFamily, StringComparison.OrdinalIgnoreCase)
                && CharacterSpacingTwips == other.CharacterSpacingTwips
                && string.Equals(Language, other.Language, StringComparison.OrdinalIgnoreCase)
                && string.Equals(EastAsiaLanguage, other.EastAsiaLanguage, StringComparison.OrdinalIgnoreCase)
                && Specified == other.Specified;
        }

        public override bool Equals(object? obj) {
            return obj is LegacyDocCharacterFormat other && Equals(other);
        }

        public override int GetHashCode() {
            int hash = 17;
            hash = (hash * 31) + Bold.GetHashCode();
            hash = (hash * 31) + Italic.GetHashCode();
            hash = (hash * 31) + Strike.GetHashCode();
            hash = (hash * 31) + DoubleStrike.GetHashCode();
            hash = (hash * 31) + Outline.GetHashCode();
            hash = (hash * 31) + Shadow.GetHashCode();
            hash = (hash * 31) + Emboss.GetHashCode();
            hash = (hash * 31) + Imprint.GetHashCode();
            hash = (hash * 31) + Hidden.GetHashCode();
            hash = (hash * 31) + NoProof.GetHashCode();
            hash = (hash * 31) + Caps.GetHashCode();
            hash = (hash * 31) + VerticalPosition.GetHashCode();
            hash = (hash * 31) + Underline.GetHashCode();
            hash = (hash * 31) + Highlight.GetHashCode();
            hash = (hash * 31) + FontSizeHalfPoints.GetHashCode();
            hash = (hash * 31) + StringComparer.OrdinalIgnoreCase.GetHashCode(ColorHex ?? string.Empty);
            hash = (hash * 31) + StringComparer.OrdinalIgnoreCase.GetHashCode(FontFamily ?? string.Empty);
            hash = (hash * 31) + CharacterSpacingTwips.GetHashCode();
            hash = (hash * 31) + StringComparer.OrdinalIgnoreCase.GetHashCode(Language ?? string.Empty);
            hash = (hash * 31) + StringComparer.OrdinalIgnoreCase.GetHashCode(EastAsiaLanguage ?? string.Empty);
            hash = (hash * 31) + Specified.GetHashCode();
            return hash;
        }
    }
}
