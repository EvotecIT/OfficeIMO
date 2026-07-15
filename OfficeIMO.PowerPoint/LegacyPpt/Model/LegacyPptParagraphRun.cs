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
            IReadOnlyList<LegacyPptTabStop>? tabStops = null) {
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
            || DefaultTabSize.HasValue || TabStops.Count != 0;
    }
}
