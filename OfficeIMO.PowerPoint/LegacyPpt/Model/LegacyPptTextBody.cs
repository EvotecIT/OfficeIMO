using System.Collections.ObjectModel;

namespace OfficeIMO.PowerPoint.LegacyPpt.Model {
    /// <summary>Represents text, its category, formatting runs, and ruler from a binary PowerPoint text box.</summary>
    public sealed class LegacyPptTextBody {
        internal LegacyPptTextBody(string text, IReadOnlyList<LegacyPptCharacterRun> characterRuns,
            IReadOnlyList<LegacyPptParagraphRun> paragraphRuns, bool hasStyleRecord,
            bool hasUnprojectedCharacterFormatting, bool hasUnprojectedParagraphFormatting,
            bool isStyleTruncated = false, LegacyPptTextType? textType = null,
            LegacyPptTextRuler? ruler = null, bool hasRulerRecord = false,
            bool isRulerTruncated = false,
            IReadOnlyList<LegacyPptTextInteraction>? interactions = null,
            bool hasStyle9Record = false,
            bool isStyle9Truncated = false,
            IReadOnlyList<LegacyPptTextField>? fields = null,
            bool hasFieldRecords = false,
            bool isFieldDataMalformed = false,
            IReadOnlyList<LegacyPptTextLanguageRun>? languageRuns = null,
            bool hasTextSpecialInfoRecord = false,
            bool hasUnprojectedTextSpecialInfo = false,
            bool isTextSpecialInfoTruncated = false) {
            Text = text ?? string.Empty;
            CharacterRuns = new ReadOnlyCollection<LegacyPptCharacterRun>(
                characterRuns?.ToArray() ?? throw new ArgumentNullException(nameof(characterRuns)));
            ParagraphRuns = new ReadOnlyCollection<LegacyPptParagraphRun>(
                paragraphRuns?.ToArray() ?? throw new ArgumentNullException(nameof(paragraphRuns)));
            HasStyleRecord = hasStyleRecord;
            HasUnprojectedCharacterFormatting = hasUnprojectedCharacterFormatting;
            HasUnprojectedParagraphFormatting = hasUnprojectedParagraphFormatting;
            IsStyleTruncated = isStyleTruncated;
            TextType = textType;
            Ruler = ruler;
            HasRulerRecord = hasRulerRecord;
            IsRulerTruncated = isRulerTruncated;
            Interactions = new ReadOnlyCollection<LegacyPptTextInteraction>(
                interactions?.ToArray() ?? Array.Empty<LegacyPptTextInteraction>());
            HasStyle9Record = hasStyle9Record;
            IsStyle9Truncated = isStyle9Truncated;
            Fields = new ReadOnlyCollection<LegacyPptTextField>(
                fields?.ToArray() ?? Array.Empty<LegacyPptTextField>());
            HasFieldRecords = hasFieldRecords;
            IsFieldDataMalformed = isFieldDataMalformed;
            LanguageRuns = new ReadOnlyCollection<LegacyPptTextLanguageRun>(
                languageRuns?.ToArray()
                    ?? Array.Empty<LegacyPptTextLanguageRun>());
            HasTextSpecialInfoRecord = hasTextSpecialInfoRecord;
            HasUnprojectedTextSpecialInfo = hasUnprojectedTextSpecialInfo;
            IsTextSpecialInfoTruncated = isTextSpecialInfoTruncated;
        }

        /// <summary>Gets the normalized text, with binary paragraph separators represented by line feeds.</summary>
        public string Text { get; }

        /// <summary>Gets character-formatting runs clipped to the exposed text.</summary>
        public IReadOnlyList<LegacyPptCharacterRun> CharacterRuns { get; }

        /// <summary>Gets paragraph-formatting runs clipped to the exposed text.</summary>
        public IReadOnlyList<LegacyPptParagraphRun> ParagraphRuns { get; }

        /// <summary>Gets whether the binary text box contains a StyleTextPropAtom.</summary>
        public bool HasStyleRecord { get; }

        /// <summary>Gets whether paragraph formatting or a nonzero paragraph level was decoded.</summary>
        public bool HasParagraphFormatting => ParagraphRuns.Any(run => run.HasExplicitFormatting)
            || Ruler?.HasFormatting == true;

        /// <summary>Gets whether decoded character formatting includes fields not yet projected natively.</summary>
        public bool HasUnprojectedCharacterFormatting { get; }

        /// <summary>Gets whether decoded paragraph formatting includes fields not yet projected natively.</summary>
        public bool HasUnprojectedParagraphFormatting { get; }

        /// <summary>Gets whether the style record was malformed or truncated.</summary>
        public bool IsStyleTruncated { get; }

        /// <summary>Gets whether the text box contains a PPT9 extended-style atom.</summary>
        public bool HasStyle9Record { get; }

        /// <summary>Gets whether the PPT9 extended-style atom is malformed or cannot be linked to its character runs.</summary>
        public bool IsStyle9Truncated { get; }

        /// <summary>Gets the text category from the TextHeaderAtom, when valid.</summary>
        public LegacyPptTextType? TextType { get; }

        /// <summary>Gets the decoded text ruler, when present and valid.</summary>
        public LegacyPptTextRuler? Ruler { get; }

        /// <summary>Gets whether the binary text box contains a TextRulerAtom.</summary>
        public bool HasRulerRecord { get; }

        /// <summary>Gets whether the text ruler was malformed or truncated.</summary>
        public bool IsRulerTruncated { get; }

        /// <summary>Gets click and mouse-over actions anchored to text ranges.</summary>
        public IReadOnlyList<LegacyPptTextInteraction> Interactions { get; }

        /// <summary>Gets native dynamic fields anchored in the text.</summary>
        public IReadOnlyList<LegacyPptTextField> Fields { get; }

        /// <summary>Gets whether the binary text box contains field metacharacter records.</summary>
        public bool HasFieldRecords { get; }

        /// <summary>Gets whether one or more field records were malformed or overlapped.</summary>
        public bool IsFieldDataMalformed { get; }

        /// <summary>
        /// Gets language and proofing runs over the logical text plus its
        /// terminal paragraph marker.
        /// </summary>
        public IReadOnlyList<LegacyPptTextLanguageRun> LanguageRuns { get; }

        /// <summary>Gets whether the binary text box contains a TextSpecialInfoAtom.</summary>
        public bool HasTextSpecialInfoRecord { get; }

        /// <summary>Gets whether the special-info atom contains fields that remain preserve-only.</summary>
        public bool HasUnprojectedTextSpecialInfo { get; }

        /// <summary>Gets whether the special-info atom is malformed or does not cover the text exactly.</summary>
        public bool IsTextSpecialInfoTruncated { get; }

        /// <summary>Gets whether at least one text range has explicit language or proofing metadata.</summary>
        public bool HasLanguageInformation => LanguageRuns.Any(run =>
            run.Language != null || run.AlternativeLanguage != null
                || run.NoProof || run.SpellingError.HasValue
                || run.NeedsRecheck.HasValue);

        /// <summary>Gets whether the text contains at least one interactive range.</summary>
        public bool HasInteractions => Interactions.Count > 0;

        /// <summary>Gets whether at least one character run carries explicit formatting.</summary>
        public bool HasExplicitCharacterFormatting => CharacterRuns.Any(run => run.HasExplicitFormatting);

        internal static LegacyPptTextBody Plain(string text) => new(text ?? string.Empty,
            Array.Empty<LegacyPptCharacterRun>(), Array.Empty<LegacyPptParagraphRun>(),
            hasStyleRecord: false, hasUnprojectedCharacterFormatting: false,
            hasUnprojectedParagraphFormatting: false);

        internal LegacyPptTextBody WithInteractions(
            IReadOnlyList<LegacyPptTextInteraction> interactions) => new(Text,
                CharacterRuns, ParagraphRuns, HasStyleRecord,
                HasUnprojectedCharacterFormatting, HasUnprojectedParagraphFormatting,
                IsStyleTruncated, TextType, Ruler, HasRulerRecord,
                IsRulerTruncated, interactions, HasStyle9Record,
                IsStyle9Truncated, Fields, HasFieldRecords,
                IsFieldDataMalformed, LanguageRuns,
                HasTextSpecialInfoRecord, HasUnprojectedTextSpecialInfo,
                IsTextSpecialInfoTruncated);

        internal LegacyPptTextBody WithFields(
            IReadOnlyList<LegacyPptTextField> fields,
            bool hasFieldRecords, bool isFieldDataMalformed) => new(Text,
                CharacterRuns, ParagraphRuns, HasStyleRecord,
                HasUnprojectedCharacterFormatting || isFieldDataMalformed,
                HasUnprojectedParagraphFormatting, IsStyleTruncated,
                TextType, Ruler, HasRulerRecord, IsRulerTruncated,
                Interactions, HasStyle9Record, IsStyle9Truncated, fields,
                hasFieldRecords, isFieldDataMalformed, LanguageRuns,
                HasTextSpecialInfoRecord, HasUnprojectedTextSpecialInfo,
                IsTextSpecialInfoTruncated);

        internal LegacyPptTextBody WithPpt9Formatting(
            IReadOnlyList<LegacyPptParagraphRun> paragraphRuns,
            bool hasUnprojectedParagraphFormatting,
            bool isStyle9Truncated) => new(Text, CharacterRuns,
                paragraphRuns, HasStyleRecord,
                HasUnprojectedCharacterFormatting,
                HasUnprojectedParagraphFormatting
                    || hasUnprojectedParagraphFormatting,
                IsStyleTruncated, TextType, Ruler, HasRulerRecord,
                IsRulerTruncated, Interactions, hasStyle9Record: true,
                isStyle9Truncated, Fields, HasFieldRecords,
                IsFieldDataMalformed, LanguageRuns,
                HasTextSpecialInfoRecord, HasUnprojectedTextSpecialInfo,
                IsTextSpecialInfoTruncated);

        internal LegacyPptTextBody WithLanguageInformation(
            IReadOnlyList<LegacyPptTextLanguageRun> languageRuns,
            bool hasTextSpecialInfoRecord,
            bool hasUnprojectedTextSpecialInfo,
            bool isTextSpecialInfoTruncated) => new(Text, CharacterRuns,
                ParagraphRuns, HasStyleRecord,
                HasUnprojectedCharacterFormatting,
                HasUnprojectedParagraphFormatting, IsStyleTruncated,
                TextType, Ruler, HasRulerRecord, IsRulerTruncated,
                Interactions, HasStyle9Record, IsStyle9Truncated, Fields,
                HasFieldRecords, IsFieldDataMalformed, languageRuns,
                hasTextSpecialInfoRecord, hasUnprojectedTextSpecialInfo,
                isTextSpecialInfoTruncated);
    }

    /// <summary>
    /// Represents language and proofing metadata for a counted binary text
    /// range. The range can include paragraph markers and the terminal marker
    /// immediately after <see cref="LegacyPptTextBody.Text"/>.
    /// </summary>
    public sealed class LegacyPptTextLanguageRun {
        internal LegacyPptTextLanguageRun(int start, int length,
            ushort? languageId, string? language,
            ushort? alternativeLanguageId, string? alternativeLanguage,
            bool noProof, bool? spellingError, bool? needsRecheck,
            bool hasUnprojectedInformation) {
            Start = start;
            Length = length;
            LanguageId = languageId;
            Language = language;
            AlternativeLanguageId = alternativeLanguageId;
            AlternativeLanguage = alternativeLanguage;
            NoProof = noProof;
            SpellingError = spellingError;
            NeedsRecheck = needsRecheck;
            HasUnprojectedInformation = hasUnprojectedInformation;
        }

        /// <summary>Gets the zero-based offset in the logical text and terminal-marker sequence.</summary>
        public int Start { get; }

        /// <summary>Gets the number of logical characters covered by this run.</summary>
        public int Length { get; }

        /// <summary>Gets the native primary language identifier, when present.</summary>
        public ushort? LanguageId { get; }

        /// <summary>Gets the primary BCP 47 language tag, when the LCID is projectable.</summary>
        public string? Language { get; }

        /// <summary>Gets the native alternate language identifier, when present.</summary>
        public ushort? AlternativeLanguageId { get; }

        /// <summary>Gets the alternate BCP 47 language tag, when the LCID is projectable.</summary>
        public string? AlternativeLanguage { get; }

        /// <summary>Gets whether native proofing is disabled for this range.</summary>
        public bool NoProof { get; }

        /// <summary>Gets whether the native spell checker marked the range as misspelled.</summary>
        public bool? SpellingError { get; }

        /// <summary>Gets whether the native spell checker marked the range for rechecking.</summary>
        public bool? NeedsRecheck { get; }

        /// <summary>Gets whether the run contains special-info fields that remain preserve-only.</summary>
        public bool HasUnprojectedInformation { get; }
    }

    /// <summary>Represents one character-formatting run from a binary PowerPoint text box.</summary>
    public sealed class LegacyPptCharacterRun {
        internal LegacyPptCharacterRun(int start, int length, string text,
            bool? bold, bool? italic, bool? underline,
            bool? shadow, bool? farEastHint, bool? kumi, bool? emboss,
            ushort? fontIndex, ushort? oldEastAsianFontIndex, ushort? ansiFontIndex,
            ushort? symbolFontIndex, string? typeface, string? oldEastAsianTypeface,
            string? ansiTypeface, string? symbolTypeface, short? fontSizePoints, string? color,
            byte? colorSchemeIndex, short? baselinePositionPercent,
            bool hasUnprojectedFormatting, byte? ppt9RunId = null) {
            Start = start;
            Length = length;
            Text = text ?? string.Empty;
            Bold = bold;
            Italic = italic;
            Underline = underline;
            Shadow = shadow;
            FarEastHint = farEastHint;
            Kumi = kumi;
            Emboss = emboss;
            FontIndex = fontIndex;
            OldEastAsianFontIndex = oldEastAsianFontIndex;
            AnsiFontIndex = ansiFontIndex;
            SymbolFontIndex = symbolFontIndex;
            Typeface = typeface;
            OldEastAsianTypeface = oldEastAsianTypeface;
            AnsiTypeface = ansiTypeface;
            SymbolTypeface = symbolTypeface;
            FontSizePoints = fontSizePoints;
            Color = color;
            ColorSchemeIndex = colorSchemeIndex;
            BaselinePositionPercent = baselinePositionPercent;
            HasUnprojectedFormatting = hasUnprojectedFormatting;
            Ppt9RunId = ppt9RunId;
        }

        /// <summary>Gets the zero-based character offset in <see cref="LegacyPptTextBody.Text"/>.</summary>
        public int Start { get; }

        /// <summary>Gets the number of exposed characters covered by the run.</summary>
        public int Length { get; }

        /// <summary>Gets the exposed text covered by the run.</summary>
        public string Text { get; }

        /// <summary>Gets the explicit bold override, or null when inherited.</summary>
        public bool? Bold { get; }

        /// <summary>Gets the explicit italic override, or null when inherited.</summary>
        public bool? Italic { get; }

        /// <summary>Gets the explicit underline override, or null when inherited.</summary>
        public bool? Underline { get; }

        /// <summary>Gets the explicit legacy character-shadow override, or null when inherited.</summary>
        public bool? Shadow { get; }

        /// <summary>Gets the explicit Far East hint override, or null when inherited.</summary>
        public bool? FarEastHint { get; }

        /// <summary>Gets the explicit kumi override, or null when inherited.</summary>
        public bool? Kumi { get; }

        /// <summary>Gets the explicit emboss override, or null when inherited.</summary>
        public bool? Emboss { get; }

        /// <summary>Gets the legacy primary typeface index, or null when inherited.</summary>
        public ushort? FontIndex { get; }

        /// <summary>Gets the legacy old East Asian typeface index, or null when inherited.</summary>
        public ushort? OldEastAsianFontIndex { get; }

        /// <summary>Gets the legacy ANSI typeface index, or null when inherited.</summary>
        public ushort? AnsiFontIndex { get; }

        /// <summary>Gets the legacy symbol typeface index, or null when inherited.</summary>
        public ushort? SymbolFontIndex { get; }

        /// <summary>Gets the resolved primary typeface name, or null when inherited or unresolved.</summary>
        public string? Typeface { get; }

        /// <summary>Gets the resolved old East Asian typeface name, or null when inherited or unresolved.</summary>
        public string? OldEastAsianTypeface { get; }

        /// <summary>Gets the resolved ANSI typeface name, or null when inherited or unresolved.</summary>
        public string? AnsiTypeface { get; }

        /// <summary>Gets the resolved symbol typeface name, or null when inherited or unresolved.</summary>
        public string? SymbolTypeface { get; }

        /// <summary>Gets the explicit font size in points, or null when inherited.</summary>
        public short? FontSizePoints { get; }

        /// <summary>Gets the resolved explicit color as RRGGBB, when available.</summary>
        public string? Color { get; }

        /// <summary>Gets the legacy scheme color index for an explicit scheme color.</summary>
        public byte? ColorSchemeIndex { get; }

        /// <summary>Gets the explicit baseline position as a percentage of line height.</summary>
        public short? BaselinePositionPercent { get; }

        /// <summary>Gets whether the run has explicit fields that are retained but not projected yet.</summary>
        public bool HasUnprojectedFormatting { get; }

        /// <summary>Gets the PPT9 extended-style grouping identifier, or null when the default group is used.</summary>
        public byte? Ppt9RunId { get; }

        /// <summary>Gets whether the run carries any explicit character-formatting field.</summary>
        public bool HasExplicitFormatting => Bold.HasValue || Italic.HasValue || Underline.HasValue
            || Shadow.HasValue || FarEastHint.HasValue || Kumi.HasValue || Emboss.HasValue
            || FontIndex.HasValue || OldEastAsianFontIndex.HasValue || AnsiFontIndex.HasValue
            || SymbolFontIndex.HasValue || Typeface != null || OldEastAsianTypeface != null
            || AnsiTypeface != null || SymbolTypeface != null || FontSizePoints.HasValue || Color != null
            || ColorSchemeIndex.HasValue || BaselinePositionPercent.HasValue;
    }
}
