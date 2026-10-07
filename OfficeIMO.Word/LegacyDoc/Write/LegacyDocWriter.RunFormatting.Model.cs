using OfficeIMO.Word.LegacyDoc.Model;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        [Flags]
        private enum LegacyDocWritableFormattingProperties {
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
            Special = 1 << 10,
            Caps = 1 << 11,
            VerticalPosition = 1 << 12,
            Underline = 1 << 13,
            Highlight = 1 << 14,
            FontSize = 1 << 15,
            Color = 1 << 16,
            FontFamily = 1 << 17,
            CharacterSpacing = 1 << 18,
            Language = 1 << 19,
            Kerning = 1 << 20,
            CharacterScale = 1 << 21
        }

        private readonly struct LegacyDocWritableFormatting : IEquatable<LegacyDocWritableFormatting> {
            internal static readonly LegacyDocWritableFormatting Plain = new LegacyDocWritableFormatting(false, false, false, false, false, false, false, false, false, false, false, null, null, null, null, null, null, null);
            internal static readonly LegacyDocWritableFormatting SpecialCharacter = new LegacyDocWritableFormatting(false, false, false, false, false, false, false, false, false, false, true, null, null, null, null, null, null, null, LegacyDocWritableFormattingProperties.Special);

            internal LegacyDocWritableFormatting(bool bold, bool italic, bool strike, bool doubleStrike, bool outline, bool shadow, bool emboss, bool imprint, bool hidden, bool noProof, bool special, byte? caps, byte? verticalPosition, byte? underline, byte? highlight, int? fontSizeHalfPoints, string? colorHex, string? fontFamily, LegacyDocWritableFormattingProperties specified = LegacyDocWritableFormattingProperties.None, int? characterSpacingTwips = null, ushort? languageId = null, ushort? eastAsiaLanguageId = null, LegacyDocRevision revision = default, int? kerningMinimumFontSizeHalfPoints = null, int? characterScalePercentage = null) {
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
                Special = special;
                Caps = caps;
                VerticalPosition = verticalPosition;
                Underline = underline;
                Highlight = highlight;
                FontSizeHalfPoints = fontSizeHalfPoints;
                ColorHex = colorHex;
                FontFamily = fontFamily;
                CharacterSpacingTwips = characterSpacingTwips;
                CharacterScalePercentage = characterScalePercentage;
                KerningMinimumFontSizeHalfPoints = kerningMinimumFontSizeHalfPoints;
                LanguageId = languageId;
                EastAsiaLanguageId = eastAsiaLanguageId;
                Specified = specified;
                Revision = revision;
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

            internal bool Special { get; }

            internal byte? Caps { get; }

            internal byte? VerticalPosition { get; }

            internal byte? Underline { get; }

            internal byte? Highlight { get; }

            internal int? FontSizeHalfPoints { get; }

            internal string? ColorHex { get; }

            internal string? FontFamily { get; }

            internal int? CharacterSpacingTwips { get; }

            internal int? CharacterScalePercentage { get; }

            internal int? KerningMinimumFontSizeHalfPoints { get; }

            internal ushort? LanguageId { get; }

            internal ushort? EastAsiaLanguageId { get; }

            internal LegacyDocRevision Revision { get; }

            internal bool HasFormatting => Bold || Italic || Strike || DoubleStrike || Outline || Shadow || Emboss || Imprint || Hidden || NoProof || Special || Caps != null || VerticalPosition != null || Underline != null || Highlight != null || FontSizeHalfPoints != null || ColorHex != null || FontFamily != null || KerningMinimumFontSizeHalfPoints != null || CharacterSpacingTwips != null || CharacterScalePercentage != null || LanguageId != null || EastAsiaLanguageId != null || Revision.HasValue || HasExplicitOffFormatting;

            private LegacyDocWritableFormattingProperties Specified { get; }

            internal LegacyDocWritableFormatting WithInheritedFormatting(LegacyDocWritableFormatting inherited) {
                if (!inherited.HasFormatting || Special) {
                    return this;
                }

                return new LegacyDocWritableFormatting(
                    IsSpecified(LegacyDocWritableFormattingProperties.Bold) ? Bold : inherited.Bold,
                    IsSpecified(LegacyDocWritableFormattingProperties.Italic) ? Italic : inherited.Italic,
                    IsSpecified(LegacyDocWritableFormattingProperties.Strike) ? Strike : inherited.Strike,
                    IsSpecified(LegacyDocWritableFormattingProperties.DoubleStrike) ? DoubleStrike : inherited.DoubleStrike,
                    IsSpecified(LegacyDocWritableFormattingProperties.Outline) ? Outline : inherited.Outline,
                    IsSpecified(LegacyDocWritableFormattingProperties.Shadow) ? Shadow : inherited.Shadow,
                    IsSpecified(LegacyDocWritableFormattingProperties.Emboss) ? Emboss : inherited.Emboss,
                    IsSpecified(LegacyDocWritableFormattingProperties.Imprint) ? Imprint : inherited.Imprint,
                    IsSpecified(LegacyDocWritableFormattingProperties.Hidden) ? Hidden : inherited.Hidden,
                    IsSpecified(LegacyDocWritableFormattingProperties.NoProof) ? NoProof : inherited.NoProof,
                    Special,
                    IsSpecified(LegacyDocWritableFormattingProperties.Caps) ? Caps : inherited.Caps,
                    IsSpecified(LegacyDocWritableFormattingProperties.VerticalPosition) ? VerticalPosition : inherited.VerticalPosition,
                    IsSpecified(LegacyDocWritableFormattingProperties.Underline) ? Underline : inherited.Underline,
                    IsSpecified(LegacyDocWritableFormattingProperties.Highlight) ? Highlight : inherited.Highlight,
                    IsSpecified(LegacyDocWritableFormattingProperties.FontSize) ? FontSizeHalfPoints : inherited.FontSizeHalfPoints,
                    IsSpecified(LegacyDocWritableFormattingProperties.Color) ? ColorHex : inherited.ColorHex,
                    IsSpecified(LegacyDocWritableFormattingProperties.FontFamily) ? FontFamily : inherited.FontFamily,
                    Specified | inherited.Specified,
                    characterSpacingTwips: IsSpecified(LegacyDocWritableFormattingProperties.CharacterSpacing) ? CharacterSpacingTwips : inherited.CharacterSpacingTwips,
                    languageId: IsSpecified(LegacyDocWritableFormattingProperties.Language) ? LanguageId : inherited.LanguageId,
                    eastAsiaLanguageId: IsSpecified(LegacyDocWritableFormattingProperties.Language) ? EastAsiaLanguageId : inherited.EastAsiaLanguageId,
                    revision: Revision.HasValue ? Revision : inherited.Revision,
                    kerningMinimumFontSizeHalfPoints: IsSpecified(LegacyDocWritableFormattingProperties.Kerning) ? KerningMinimumFontSizeHalfPoints : inherited.KerningMinimumFontSizeHalfPoints,
                    characterScalePercentage: IsSpecified(LegacyDocWritableFormattingProperties.CharacterScale) ? CharacterScalePercentage : inherited.CharacterScalePercentage);
            }

            // Note reference characters retain source typography as well as the
            // binary special-character flag used to identify their marker.
            internal LegacyDocWritableFormatting WithSpecialCharacter() => new(
                Bold, Italic, Strike, DoubleStrike, Outline, Shadow, Emboss, Imprint, Hidden, NoProof,
                true, Caps, VerticalPosition, Underline, Highlight, FontSizeHalfPoints, ColorHex, FontFamily,
                Specified | LegacyDocWritableFormattingProperties.Special, CharacterSpacingTwips,
                LanguageId, EastAsiaLanguageId, Revision, KerningMinimumFontSizeHalfPoints, CharacterScalePercentage);

            internal LegacyDocWritableFormatting WithRevision(LegacyDocRevision revision) {
                return new LegacyDocWritableFormatting(
                    Bold,
                    Italic,
                    Strike,
                    DoubleStrike,
                    Outline,
                    Shadow,
                    Emboss,
                    Imprint,
                    Hidden,
                    NoProof,
                    Special,
                    Caps,
                    VerticalPosition,
                    Underline,
                    Highlight,
                    FontSizeHalfPoints,
                    ColorHex,
                    FontFamily,
                    Specified,
                    CharacterSpacingTwips,
                    LanguageId,
                    EastAsiaLanguageId,
                    revision,
                    KerningMinimumFontSizeHalfPoints,
                    CharacterScalePercentage);
            }

            private bool HasExplicitOffFormatting =>
                (IsSpecified(LegacyDocWritableFormattingProperties.Bold) && !Bold)
                || (IsSpecified(LegacyDocWritableFormattingProperties.Italic) && !Italic)
                || (IsSpecified(LegacyDocWritableFormattingProperties.Strike) && !Strike)
                || (IsSpecified(LegacyDocWritableFormattingProperties.DoubleStrike) && !DoubleStrike)
                || (IsSpecified(LegacyDocWritableFormattingProperties.Outline) && !Outline)
                || (IsSpecified(LegacyDocWritableFormattingProperties.Shadow) && !Shadow)
                || (IsSpecified(LegacyDocWritableFormattingProperties.Emboss) && !Emboss)
                || (IsSpecified(LegacyDocWritableFormattingProperties.Imprint) && !Imprint)
                || (IsSpecified(LegacyDocWritableFormattingProperties.Hidden) && !Hidden)
                || (IsSpecified(LegacyDocWritableFormattingProperties.NoProof) && !NoProof)
                || (IsSpecified(LegacyDocWritableFormattingProperties.Caps) && Caps == null)
                || (IsSpecified(LegacyDocWritableFormattingProperties.VerticalPosition) && VerticalPosition == null)
                || (IsSpecified(LegacyDocWritableFormattingProperties.Underline) && Underline == null)
                || (IsSpecified(LegacyDocWritableFormattingProperties.Highlight) && Highlight == null)
                || (IsSpecified(LegacyDocWritableFormattingProperties.CharacterSpacing) && CharacterSpacingTwips == null)
                || (IsSpecified(LegacyDocWritableFormattingProperties.Language) && LanguageId == null && EastAsiaLanguageId == null);

            internal bool IsSpecified(LegacyDocWritableFormattingProperties property) {
                return (Specified & property) != 0;
            }

            // Adjacent native records must retain explicit overrides even when
            // their values match. Field uniformity compares effective values instead.
            internal bool HasSameEncoding(LegacyDocWritableFormatting other) =>
                Specified == other.Specified && Equals(other);

            public bool Equals(LegacyDocWritableFormatting other) {
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
                    && Special == other.Special
                    && Caps == other.Caps
                    && VerticalPosition == other.VerticalPosition
                    && Underline == other.Underline
                    && Highlight == other.Highlight
                    && FontSizeHalfPoints == other.FontSizeHalfPoints
                    && string.Equals(ColorHex, other.ColorHex, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(FontFamily, other.FontFamily, StringComparison.OrdinalIgnoreCase)
                    && CharacterSpacingTwips == other.CharacterSpacingTwips
                    && CharacterScalePercentage == other.CharacterScalePercentage
                    && KerningMinimumFontSizeHalfPoints == other.KerningMinimumFontSizeHalfPoints
                    && LanguageId == other.LanguageId
                    && EastAsiaLanguageId == other.EastAsiaLanguageId
                    && Revision.Equals(other.Revision);
            }

            public override bool Equals(object? obj) {
                return obj is LegacyDocWritableFormatting other && Equals(other);
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
                hash = (hash * 31) + Special.GetHashCode();
                hash = (hash * 31) + Caps.GetHashCode();
                hash = (hash * 31) + VerticalPosition.GetHashCode();
                hash = (hash * 31) + Underline.GetHashCode();
                hash = (hash * 31) + Highlight.GetHashCode();
                hash = (hash * 31) + FontSizeHalfPoints.GetHashCode();
                hash = (hash * 31) + StringComparer.OrdinalIgnoreCase.GetHashCode(ColorHex ?? string.Empty);
                hash = (hash * 31) + StringComparer.OrdinalIgnoreCase.GetHashCode(FontFamily ?? string.Empty);
                hash = (hash * 31) + CharacterSpacingTwips.GetHashCode();
                hash = (hash * 31) + CharacterScalePercentage.GetHashCode();
                hash = (hash * 31) + KerningMinimumFontSizeHalfPoints.GetHashCode();
                hash = (hash * 31) + LanguageId.GetHashCode();
                hash = (hash * 31) + EastAsiaLanguageId.GetHashCode();
                hash = (hash * 31) + Revision.GetHashCode();
                return hash;
            }
        }

        private readonly struct LegacyDocWritableLanguageIds {
            internal LegacyDocWritableLanguageIds(ushort? languageId, ushort? eastAsiaLanguageId) {
                LanguageId = languageId;
                EastAsiaLanguageId = eastAsiaLanguageId;
            }

            internal ushort? LanguageId { get; }

            internal ushort? EastAsiaLanguageId { get; }

            internal bool HasAny => LanguageId != null || EastAsiaLanguageId != null;
        }

        private readonly struct LegacyDocWritableRun {
            internal LegacyDocWritableRun(int startCharacter, int length, LegacyDocWritableFormatting formatting, int? pictureDataOffset = null) {
                StartCharacter = startCharacter;
                Length = length;
                Formatting = formatting;
                PictureDataOffset = pictureDataOffset;
            }

            internal int StartCharacter { get; }

            internal int Length { get; }

            internal int EndCharacter => StartCharacter + Length;

            internal LegacyDocWritableFormatting Formatting { get; }

            internal int? PictureDataOffset { get; }

            internal LegacyDocWritableRun Extend(int additionalLength) {
                return new LegacyDocWritableRun(StartCharacter, Length + additionalLength, Formatting, PictureDataOffset);
            }
        }

        private readonly struct LegacyDocWritableSegment {
            internal LegacyDocWritableSegment(int startCharacter, int length, LegacyDocWritableFormatting formatting, int? pictureDataOffset = null) {
                StartCharacter = startCharacter;
                Length = length;
                Formatting = formatting;
                PictureDataOffset = pictureDataOffset;
            }

            internal int StartCharacter { get; }

            internal int Length { get; }

            internal int EndCharacter => StartCharacter + Length;

            internal LegacyDocWritableFormatting Formatting { get; }

            internal int? PictureDataOffset { get; }

            internal bool HasFormatting => Formatting.HasFormatting || PictureDataOffset != null;

            internal LegacyDocWritableSegment Extend(int additionalLength) {
                return new LegacyDocWritableSegment(StartCharacter, Length + additionalLength, Formatting, PictureDataOffset);
            }
        }
    }
}
