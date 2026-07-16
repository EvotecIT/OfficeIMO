namespace OfficeIMO.PowerPoint.LegacyPpt.Model {
    /// <summary>Specifies a binary PowerPoint paragraph alignment.</summary>
    public enum LegacyPptTextAlignment : ushort {
        /// <summary>Left for horizontal text or top for vertical text.</summary>
        Left = 0,
        /// <summary>Centered.</summary>
        Center = 1,
        /// <summary>Right for horizontal text or bottom for vertical text.</summary>
        Right = 2,
        /// <summary>Flush left and right.</summary>
        Justify = 3,
        /// <summary>Distributed character spacing.</summary>
        Distributed = 4,
        /// <summary>Thai distributed alignment.</summary>
        ThaiDistributed = 5,
        /// <summary>Low Kashida justification.</summary>
        JustifyLow = 6
    }

    /// <summary>Specifies how characters align within a line height.</summary>
    public enum LegacyPptFontAlignment : ushort {
        /// <summary>Align on the font baseline.</summary>
        Baseline = 0,
        /// <summary>Hang characters from the top of the line.</summary>
        Hanging = 1,
        /// <summary>Center characters within the line height.</summary>
        Center = 2,
        /// <summary>Anchor characters to the bottom of the line.</summary>
        Bottom = 3
    }

    /// <summary>Specifies paragraph text direction.</summary>
    public enum LegacyPptTextDirection : ushort {
        /// <summary>Left-to-right text.</summary>
        LeftToRight = 0,
        /// <summary>Right-to-left text.</summary>
        RightToLeft = 1
    }

    /// <summary>Specifies a classic PowerPoint automatic-numbering scheme.</summary>
    public enum LegacyPptAutoNumberScheme : ushort {
        /// <summary>Lowercase Latin letter followed by a period.</summary>
        AlphaLowerPeriod = 0x0000,
        /// <summary>Uppercase Latin letter followed by a period.</summary>
        AlphaUpperPeriod = 0x0001,
        /// <summary>Arabic numeral followed by a closing parenthesis.</summary>
        ArabicParenRight = 0x0002,
        /// <summary>Arabic numeral followed by a period.</summary>
        ArabicPeriod = 0x0003,
        /// <summary>Lowercase Roman numeral enclosed in parentheses.</summary>
        RomanLowerParenBoth = 0x0004,
        /// <summary>Lowercase Roman numeral followed by a closing parenthesis.</summary>
        RomanLowerParenRight = 0x0005,
        /// <summary>Lowercase Roman numeral followed by a period.</summary>
        RomanLowerPeriod = 0x0006,
        /// <summary>Uppercase Roman numeral followed by a period.</summary>
        RomanUpperPeriod = 0x0007,
        /// <summary>Lowercase Latin letter enclosed in parentheses.</summary>
        AlphaLowerParenBoth = 0x0008,
        /// <summary>Lowercase Latin letter followed by a closing parenthesis.</summary>
        AlphaLowerParenRight = 0x0009,
        /// <summary>Uppercase Latin letter enclosed in parentheses.</summary>
        AlphaUpperParenBoth = 0x000A,
        /// <summary>Uppercase Latin letter followed by a closing parenthesis.</summary>
        AlphaUpperParenRight = 0x000B,
        /// <summary>Arabic numeral enclosed in parentheses.</summary>
        ArabicParenBoth = 0x000C,
        /// <summary>Arabic numeral without a delimiter.</summary>
        ArabicPlain = 0x000D,
        /// <summary>Uppercase Roman numeral enclosed in parentheses.</summary>
        RomanUpperParenBoth = 0x000E,
        /// <summary>Uppercase Roman numeral followed by a closing parenthesis.</summary>
        RomanUpperParenRight = 0x000F,
        /// <summary>Simplified Chinese numbering without a delimiter.</summary>
        SimplifiedChinesePlain = 0x0010,
        /// <summary>Simplified Chinese numbering followed by a period.</summary>
        SimplifiedChinesePeriod = 0x0011,
        /// <summary>Double-byte circled numbers.</summary>
        CircleNumberDoubleBytePlain = 0x0012,
        /// <summary>Wingdings white circled numbers.</summary>
        CircleNumberWingdingsWhitePlain = 0x0013,
        /// <summary>Wingdings black circled numbers.</summary>
        CircleNumberWingdingsBlackPlain = 0x0014,
        /// <summary>Traditional Chinese numbering without a delimiter.</summary>
        TraditionalChinesePlain = 0x0015,
        /// <summary>Traditional Chinese numbering followed by a period.</summary>
        TraditionalChinesePeriod = 0x0016,
        /// <summary>Bidirectional Arabic alphabetic numbering followed by a minus.</summary>
        Arabic1Minus = 0x0017,
        /// <summary>Bidirectional Arabic abjad numbering followed by a minus.</summary>
        Arabic2Minus = 0x0018,
        /// <summary>Bidirectional Hebrew numbering followed by a minus.</summary>
        Hebrew2Minus = 0x0019,
        /// <summary>Japanese or Korean numbering without a delimiter.</summary>
        JapaneseKoreanPlain = 0x001A,
        /// <summary>Japanese or Korean numbering followed by a period.</summary>
        JapaneseKoreanPeriod = 0x001B,
        /// <summary>Double-byte Arabic numbering without a delimiter.</summary>
        ArabicDoubleBytePlain = 0x001C,
        /// <summary>Double-byte Arabic numbering followed by a period.</summary>
        ArabicDoubleBytePeriod = 0x001D,
        /// <summary>Thai alphabetic numbering followed by a period.</summary>
        ThaiAlphaPeriod = 0x001E,
        /// <summary>Thai alphabetic numbering followed by a closing parenthesis.</summary>
        ThaiAlphaParenRight = 0x001F,
        /// <summary>Thai alphabetic numbering enclosed in parentheses.</summary>
        ThaiAlphaParenBoth = 0x0020,
        /// <summary>Thai numeric numbering followed by a period.</summary>
        ThaiNumberPeriod = 0x0021,
        /// <summary>Thai numeric numbering followed by a closing parenthesis.</summary>
        ThaiNumberParenRight = 0x0022,
        /// <summary>Thai numeric numbering enclosed in parentheses.</summary>
        ThaiNumberParenBoth = 0x0023,
        /// <summary>Hindi alphabetic numbering followed by a period.</summary>
        HindiAlphaPeriod = 0x0024,
        /// <summary>Hindi numeric numbering followed by a period.</summary>
        HindiNumberPeriod = 0x0025,
        /// <summary>Japanese numbering followed by a double-byte period.</summary>
        JapaneseDoubleBytePeriod = 0x0026,
        /// <summary>Hindi numeric numbering followed by a closing parenthesis.</summary>
        HindiNumberParenRight = 0x0027,
        /// <summary>Alternate Hindi alphabetic numbering followed by a period.</summary>
        HindiAlpha1Period = 0x0028
    }

    /// <summary>Specifies how text aligns at a binary PowerPoint tab stop.</summary>
    public enum LegacyPptTabAlignment : ushort {
        /// <summary>Text starts at the tab stop.</summary>
        Left = 0,
        /// <summary>Text is centered on the tab stop.</summary>
        Center = 1,
        /// <summary>Text ends at the tab stop.</summary>
        Right = 2,
        /// <summary>Text is aligned on its decimal separator.</summary>
        Decimal = 3
    }

    /// <summary>Represents one tab stop from binary PowerPoint paragraph formatting.</summary>
    public sealed class LegacyPptTabStop {
        internal LegacyPptTabStop(short position, LegacyPptTabAlignment alignment) {
            Position = position;
            Alignment = alignment;
        }

        /// <summary>Gets the tab position in PowerPoint master units.</summary>
        public short Position { get; }

        /// <summary>Gets how text aligns at the tab position.</summary>
        public LegacyPptTabAlignment Alignment { get; }
    }

    /// <summary>Represents one paragraph-formatting run from a binary PowerPoint text box.</summary>
    public sealed class LegacyPptParagraphRun {
        internal LegacyPptParagraphRun(int start, int length, ushort indentLevel,
            bool? hasBullet, bool? bulletHasFont, bool? bulletHasColor, bool? bulletHasSize,
            char? bulletCharacter, ushort? bulletFontIndex, string? bulletTypeface,
            short? bulletSize, string? bulletColor, byte? bulletColorSchemeIndex,
            LegacyPptTextAlignment? alignment, short? lineSpacing, short? spaceBefore,
            short? spaceAfter, LegacyPptFontAlignment? fontAlignment,
            bool? characterWrap, bool? wordWrap, bool? overflow,
            LegacyPptTextDirection? textDirection, bool hasUnprojectedFormatting,
            short? leftMargin = null, short? indent = null, short? defaultTabSize = null,
            IReadOnlyList<LegacyPptTabStop>? tabStops = null,
            bool? hasAutoNumber = null,
            LegacyPptAutoNumberScheme? autoNumberScheme = null,
            short? autoNumberStartAt = null,
            ushort? bulletPictureReference = null,
            LegacyPptPictureBullet? pictureBullet = null) {
            Start = start;
            Length = length;
            IndentLevel = indentLevel;
            HasBullet = hasBullet;
            BulletHasFont = bulletHasFont;
            BulletHasColor = bulletHasColor;
            BulletHasSize = bulletHasSize;
            BulletCharacter = bulletCharacter;
            BulletFontIndex = bulletFontIndex;
            BulletTypeface = bulletTypeface;
            BulletSize = bulletSize;
            BulletColor = bulletColor;
            BulletColorSchemeIndex = bulletColorSchemeIndex;
            Alignment = alignment;
            LineSpacing = lineSpacing;
            SpaceBefore = spaceBefore;
            SpaceAfter = spaceAfter;
            FontAlignment = fontAlignment;
            CharacterWrap = characterWrap;
            WordWrap = wordWrap;
            Overflow = overflow;
            TextDirection = textDirection;
            HasUnprojectedFormatting = hasUnprojectedFormatting;
            LeftMargin = leftMargin;
            Indent = indent;
            DefaultTabSize = defaultTabSize;
            TabStops = tabStops?.ToArray() ?? Array.Empty<LegacyPptTabStop>();
            HasAutoNumber = hasAutoNumber;
            AutoNumberScheme = autoNumberScheme;
            AutoNumberStartAt = autoNumberStartAt;
            BulletPictureReference = bulletPictureReference;
            PictureBullet = pictureBullet;
        }

        /// <summary>Gets the zero-based character offset covered by this run.</summary>
        public int Start { get; }

        /// <summary>Gets the exposed character count covered by this run.</summary>
        public int Length { get; }

        /// <summary>Gets the paragraph indentation level.</summary>
        public ushort IndentLevel { get; }

        /// <summary>Gets the explicit bullet-enabled override, or null when inherited.</summary>
        public bool? HasBullet { get; }

        /// <summary>Gets whether an explicit bullet font is active.</summary>
        public bool? BulletHasFont { get; }

        /// <summary>Gets whether an explicit bullet color is active.</summary>
        public bool? BulletHasColor { get; }

        /// <summary>Gets whether an explicit bullet size is active.</summary>
        public bool? BulletHasSize { get; }

        /// <summary>Gets the explicit bullet character.</summary>
        public char? BulletCharacter { get; }

        /// <summary>Gets the explicit bullet font index.</summary>
        public ushort? BulletFontIndex { get; }

        /// <summary>Gets the resolved explicit bullet typeface.</summary>
        public string? BulletTypeface { get; }

        /// <summary>Gets the raw bullet size: positive is percent, negative is points.</summary>
        public short? BulletSize { get; }

        /// <summary>Gets the resolved explicit bullet color as RRGGBB.</summary>
        public string? BulletColor { get; }

        /// <summary>Gets the explicit legacy bullet scheme-color index.</summary>
        public byte? BulletColorSchemeIndex { get; }

        /// <summary>Gets the explicit paragraph alignment.</summary>
        public LegacyPptTextAlignment? Alignment { get; }

        /// <summary>Gets raw line spacing: nonnegative is percent, negative is master units.</summary>
        public short? LineSpacing { get; }

        /// <summary>Gets raw spacing before: nonnegative is percent, negative is master units.</summary>
        public short? SpaceBefore { get; }

        /// <summary>Gets raw spacing after: nonnegative is percent, negative is master units.</summary>
        public short? SpaceAfter { get; }

        /// <summary>Gets the explicit font alignment within the line height.</summary>
        public LegacyPptFontAlignment? FontAlignment { get; }

        /// <summary>Gets the explicit East Asian character-wrap setting.</summary>
        public bool? CharacterWrap { get; }

        /// <summary>Gets the explicit word-wrap setting.</summary>
        public bool? WordWrap { get; }

        /// <summary>Gets the explicit hanging-punctuation setting.</summary>
        public bool? Overflow { get; }

        /// <summary>Gets the explicit paragraph text direction.</summary>
        public LegacyPptTextDirection? TextDirection { get; }

        /// <summary>Gets the explicit left margin in PowerPoint master units.</summary>
        public short? LeftMargin { get; }

        /// <summary>Gets the explicit first-line indentation in PowerPoint master units.</summary>
        public short? Indent { get; }

        /// <summary>Gets the explicit default tab size in PowerPoint master units.</summary>
        public short? DefaultTabSize { get; }

        /// <summary>Gets explicit paragraph tab stops.</summary>
        public IReadOnlyList<LegacyPptTabStop> TabStops { get; }

        /// <summary>Gets whether automatic numbering is explicitly enabled, disabled, or inherited.</summary>
        public bool? HasAutoNumber { get; }

        /// <summary>Gets the automatic-numbering scheme, when explicitly present or defaulted.</summary>
        public LegacyPptAutoNumberScheme? AutoNumberScheme { get; }

        /// <summary>Gets the automatic-numbering start value, when numbering is enabled.</summary>
        public short? AutoNumberStartAt { get; }

        /// <summary>Gets the zero-based picture-bullet reference from the PPT9 text properties.</summary>
        public ushort? BulletPictureReference { get; }

        /// <summary>Gets the resolved PPT9 picture bullet, when its BLIP is importable.</summary>
        public LegacyPptPictureBullet? PictureBullet { get; }

        /// <summary>Gets whether the run includes fields that are retained but not projected yet.</summary>
        public bool HasUnprojectedFormatting { get; }

        /// <summary>Gets whether the run contains explicit paragraph formatting.</summary>
        public bool HasExplicitFormatting => IndentLevel != 0 || HasBullet.HasValue
            || BulletHasFont.HasValue || BulletHasColor.HasValue || BulletHasSize.HasValue
            || BulletCharacter.HasValue || BulletFontIndex.HasValue || BulletSize.HasValue
            || BulletColor != null || Alignment.HasValue || LineSpacing.HasValue
            || SpaceBefore.HasValue || SpaceAfter.HasValue || FontAlignment.HasValue
            || CharacterWrap.HasValue || WordWrap.HasValue || Overflow.HasValue
            || TextDirection.HasValue || LeftMargin.HasValue || Indent.HasValue
            || DefaultTabSize.HasValue || TabStops.Count != 0
            || HasAutoNumber.HasValue || AutoNumberScheme.HasValue
            || AutoNumberStartAt.HasValue || BulletPictureReference.HasValue
            || PictureBullet != null;

        internal LegacyPptParagraphRun WithPpt9Formatting(
            bool? hasAutoNumber,
            LegacyPptAutoNumberScheme? autoNumberScheme,
            short? autoNumberStartAt,
            ushort? bulletPictureReference,
            LegacyPptPictureBullet? pictureBullet,
            bool hasUnprojectedFormatting) => new(Start, Length,
                IndentLevel, HasBullet, BulletHasFont, BulletHasColor,
                BulletHasSize, BulletCharacter, BulletFontIndex,
                BulletTypeface, BulletSize, BulletColor,
                BulletColorSchemeIndex, Alignment, LineSpacing, SpaceBefore,
                SpaceAfter, FontAlignment, CharacterWrap, WordWrap, Overflow,
                TextDirection,
                HasUnprojectedFormatting || hasUnprojectedFormatting,
                LeftMargin, Indent, DefaultTabSize, TabStops, hasAutoNumber,
                autoNumberScheme, autoNumberStartAt, bulletPictureReference,
                pictureBullet);
    }
}
