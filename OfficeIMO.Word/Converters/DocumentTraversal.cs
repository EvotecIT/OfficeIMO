using DocumentFormat.OpenXml.Wordprocessing;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;

namespace OfficeIMO.Word {
    /// <summary>
    /// Provides helper methods for traversing documents and resolving list markers.
    /// </summary>
    public static class DocumentTraversal {
        /// <summary>
        /// Describes list information for a paragraph.
        /// </summary>
        public readonly struct ListInfo {
            /// <summary>
            /// Initializes a new instance of the <see cref="ListInfo"/> struct.
            /// </summary>
            /// <param name="level">Zero-based numbering level.</param>
            /// <param name="ordered"><c>true</c> if the list uses numbering; otherwise, <c>false</c>.</param>
            /// <param name="start">Starting index for the list.</param>
            /// <param name="format">Numbering format for the list.</param>
            /// <param name="text">Raw text pattern defining the marker.</param>
            /// <param name="leftIndentTwips">List text position in twentieths of a point, when defined.</param>
            /// <param name="hangingIndentTwips">List marker hanging indentation in twentieths of a point, when defined.</param>
            /// <param name="markerFontFamily">Marker font family from the numbering level, when defined.</param>
            /// <param name="markerBold">Marker bold setting from the numbering level, when defined.</param>
            /// <param name="markerItalic">Marker italic setting from the numbering level, when defined.</param>
            /// <param name="markerColorHex">Marker color from the numbering level, when defined.</param>
            /// <param name="levelJustification">Marker justification from the numbering level, when defined.</param>
            /// <param name="levelSuffix">Marker suffix from the numbering level, when defined.</param>
            public ListInfo(int level, bool ordered, int start, NumberFormatValues? format, string? text, int? leftIndentTwips = null, int? hangingIndentTwips = null, string? markerFontFamily = null, bool? markerBold = null, bool? markerItalic = null, string? markerColorHex = null, LevelJustificationValues? levelJustification = null, LevelSuffixValues? levelSuffix = null)
                : this(level, ordered, start, format, text, leftIndentTwips, hangingIndentTwips, markerFontFamily, markerBold, markerItalic, markerColorHex, markerFontSize: null, levelJustification, levelSuffix) {
            }

            internal ListInfo(int level, bool ordered, int start, NumberFormatValues? format, string? text, int? leftIndentTwips, int? hangingIndentTwips, string? markerFontFamily, bool? markerBold, bool? markerItalic, string? markerColorHex, double? markerFontSize, LevelJustificationValues? levelJustification, LevelSuffixValues? levelSuffix) {
                Level = level;
                Ordered = ordered;
                Start = start;
                NumberFormat = format;
                LevelText = text;
                LeftIndentTwips = leftIndentTwips;
                HangingIndentTwips = hangingIndentTwips;
                MarkerFontFamily = markerFontFamily;
                MarkerBold = markerBold;
                MarkerItalic = markerItalic;
                MarkerColorHex = markerColorHex;
                MarkerFontSize = markerFontSize;
                LevelJustification = levelJustification;
                LevelSuffix = levelSuffix;
            }

            /// <summary>Zero-based nesting level.</summary>
            public int Level { get; }
            /// <summary>Indicates whether numbering is used.</summary>
            public bool Ordered { get; }
            /// <summary>Starting index for the list.</summary>
            public int Start { get; }
            /// <summary>Numbering format applied to the list.</summary>
            public NumberFormatValues? NumberFormat { get; }
            /// <summary>Pattern used to build the list marker.</summary>
            public string? LevelText { get; }
            /// <summary>List text position in twentieths of a point, when defined.</summary>
            public int? LeftIndentTwips { get; }
            /// <summary>List marker hanging indentation in twentieths of a point, when defined.</summary>
            public int? HangingIndentTwips { get; }
            /// <summary>Marker font family from the numbering level, when defined.</summary>
            public string? MarkerFontFamily { get; }
            /// <summary>Marker bold setting from the numbering level, when defined.</summary>
            public bool? MarkerBold { get; }
            /// <summary>Marker italic setting from the numbering level, when defined.</summary>
            public bool? MarkerItalic { get; }
            /// <summary>Marker color from the numbering level, when defined.</summary>
            public string? MarkerColorHex { get; }
            /// <summary>Marker font size from the numbering level, in points, when defined.</summary>
            public double? MarkerFontSize { get; }
            /// <summary>Marker justification from the numbering level, when defined.</summary>
            public LevelJustificationValues? LevelJustification { get; }
            /// <summary>Marker suffix from the numbering level, when defined.</summary>
            public LevelSuffixValues? LevelSuffix { get; }
        }

        /// <summary>
        /// Enumerates all sections within the document.
        /// </summary>
        public static IEnumerable<WordSection> EnumerateSections(WordDocument document) {
            return document?.Sections ?? Enumerable.Empty<WordSection>();
        }

        /// <summary>
        /// Resolves list information for the given paragraph.
        /// </summary>
        /// <param name="paragraph">Paragraph to inspect.</param>
        /// <returns>List info for the paragraph or null when paragraph isn't a list item.</returns>
        public static ListInfo? GetListInfo(WordParagraph paragraph) {
            if (paragraph == null || !paragraph.IsListItem) {
                return null;
            }

            Dictionary<int, ListNumberingDefinition> definitions = BuildListNumberingDefinitions(paragraph._document);
            return GetListInfo(paragraph, definitions);
        }

        private static ListInfo? GetListInfo(WordParagraph paragraph, IReadOnlyDictionary<int, ListNumberingDefinition> definitions) {
            if (paragraph == null || !paragraph.IsListItem) {
                return null;
            }

            int level = paragraph.ListItemLevel ?? 0;
            int? overrideStart = null;
            int start = 1;
            NumberFormatValues? numberFormat = null;
            string? levelText = null;
            int? leftIndentTwips = null;
            int? hangingIndentTwips = null;
            string? markerFontFamily = null;
            bool? markerBold = null;
            bool? markerItalic = null;
            string? markerColorHex = null;
            double? markerFontSize = null;
            LevelJustificationValues? levelJustification = null;
            LevelSuffixValues? levelSuffix = null;

            int? numberId = paragraph._listNumberId;
            ListNumberingDefinition? definition = null;
            if (numberId.HasValue) {
                definitions.TryGetValue(numberId.Value, out definition);
            }
            if (definition != null &&
                definition.StartOverrides.TryGetValue(level, out int overrideValue) &&
                (!definition.OverridesAreDefault || overrideValue != 1)) {
                overrideStart = overrideValue;
                start = overrideValue;
            }

            if (definition != null && definition.Levels.TryGetValue(level, out ListLevelDefinition levelDefinition)) {
                if (!overrideStart.HasValue) {
                    start = levelDefinition.Start;
                }
                numberFormat = levelDefinition.NumberFormat;
                levelText = levelDefinition.LevelText;
                leftIndentTwips = levelDefinition.LeftIndentTwips;
                hangingIndentTwips = levelDefinition.HangingIndentTwips;
                markerFontFamily = levelDefinition.MarkerFontFamily;
                markerBold = levelDefinition.MarkerBold;
                markerItalic = levelDefinition.MarkerItalic;
                markerColorHex = levelDefinition.MarkerColorHex;
                markerFontSize = levelDefinition.MarkerFontSize;
                levelJustification = levelDefinition.LevelJustification;
                levelSuffix = levelDefinition.LevelSuffix;
            }

            bool ordered = definition?.Style switch {
                WordListStyle.Bulleted => false,
                WordListStyle.BulletedChars => false,
                _ => true,
            };
            return new ListInfo(level, ordered, start, numberFormat, levelText, leftIndentTwips, hangingIndentTwips, markerFontFamily, markerBold, markerItalic, markerColorHex, markerFontSize, levelJustification, levelSuffix);
        }

        private static int? ParseOptionalInt32(string? value) {
            return int.TryParse(value, out int parsed) ? parsed : null;
        }

        /// <summary>
        /// Builds a lookup of list markers for all paragraphs in the document.
        /// </summary>
        public static Dictionary<WordParagraph, (int Level, string Marker)> BuildListMarkers(WordDocument document) {
            Dictionary<WordParagraph, (int, string)> result = new(ParagraphReferenceComparer.Instance);
            Dictionary<int, ListNumberingDefinition> definitions = BuildListNumberingDefinitions(document);
            Dictionary<int, List<WordParagraph>> itemsByNumberId = BuildListItemsByNumberId(document);

            foreach (KeyValuePair<int, List<WordParagraph>> listItems in itemsByNumberId) {
                Dictionary<int, int> indices = new();
                Dictionary<int, NumberFormatValues?> formats = new();
                int lastLevel = 0;
                bool first = true;
                foreach (WordParagraph item in listItems.Value) {
                    ListInfo? info = GetListInfo(item, definitions);
                    if (info == null) {
                        continue;
                    }

                    int level = info.Value.Level;
                    if (first) {
                        lastLevel = level;
                        first = false;
                    }

                    if (level < lastLevel) {
                        foreach (int key in indices.Keys.Where(key => key > level).ToList()) {
                            indices.Remove(key);
                            formats.Remove(key);
                        }
                    }

                    lastLevel = level;
                    if (!indices.ContainsKey(level)) {
                        indices[level] = info.Value.Start;
                        formats[level] = info.Value.NumberFormat;
                    }

                    int currentIndex = indices[level];
                    indices[level] = currentIndex + 1;

                    string marker = info.Value.Ordered
                        ? BuildMarker(level, currentIndex, indices, formats, info.Value.LevelText)
                        : (info.Value.LevelText ?? "•");

                    result[item] = (level, marker);
                }
            }

            return result;
        }

        /// <summary>
        /// Builds a lookup of list numeric indices for all paragraphs in the document.
        /// The returned index is the 1-based number of the item at its nesting level,
        /// accounting for list continuation across unrelated content.
        /// </summary>
        public static Dictionary<WordParagraph, (int Level, int Index)> BuildListIndices(WordDocument document) {
            Dictionary<WordParagraph, (int, int)> result = new(ParagraphReferenceComparer.Instance);
            Dictionary<int, ListNumberingDefinition> definitions = BuildListNumberingDefinitions(document);
            Dictionary<int, List<WordParagraph>> itemsByNumberId = BuildListItemsByNumberId(document);

            foreach (KeyValuePair<int, List<WordParagraph>> listItems in itemsByNumberId) {
                // Track current numbering per level within this list
                Dictionary<int, int> indices = new();
                int lastLevel = 0;
                bool first = true;
                foreach (WordParagraph item in listItems.Value) {
                    ListInfo? info = GetListInfo(item, definitions);
                    if (info == null) continue;

                    int level = info.Value.Level;
                    if (first) { lastLevel = level; first = false; }
                    // If we moved to a shallower level, clear deeper counters so sublists restart
                    if (level < lastLevel) {
                        foreach (var key in indices.Keys.Where(k => k > level).ToList()) indices.Remove(key);
                    }
                    lastLevel = level;

                    if (!indices.ContainsKey(level)) {
                        indices[level] = info.Value.Start;
                    }

                    int currentIndex = indices[level];
                    result[item] = (level, currentIndex);
                    indices[level] = currentIndex + 1;
                }
            }

            return result;
        }

        private static Dictionary<int, ListNumberingDefinition> BuildListNumberingDefinitions(WordDocument? document) {
            var result = new Dictionary<int, ListNumberingDefinition>();
            Numbering? numbering = document?._wordprocessingDocument.MainDocumentPart?.NumberingDefinitionsPart?.Numbering;
            if (numbering == null) {
                return result;
            }

            var abstracts = new Dictionary<int, AbstractNum>();
            foreach (AbstractNum abstractNum in numbering.Elements<AbstractNum>()) {
                if (abstractNum.AbstractNumberId?.HasValue == true && !abstracts.ContainsKey(abstractNum.AbstractNumberId.Value)) {
                    abstracts.Add(abstractNum.AbstractNumberId.Value, abstractNum);
                }
            }

            foreach (NumberingInstance instance in numbering.Elements<NumberingInstance>()) {
                if (instance.NumberID?.Value == null) {
                    continue;
                }

                int numberId = instance.NumberID.Value;
                int? abstractId = instance.AbstractNumId?.Val?.Value != null ? instance.AbstractNumId.Val.Value : null;
                AbstractNum? abstractNum = abstractId.HasValue && abstracts.TryGetValue(abstractId.Value, out AbstractNum? foundAbstract)
                    ? foundAbstract
                    : null;

                var overrides = instance.Elements<LevelOverride>().ToList();
                bool overridesAreDefault = overrides.Count >= 9 &&
                    overrides.All(levelOverride => {
                        var startOverrideValue = levelOverride.GetFirstChild<StartOverrideNumberingValue>();
                        return startOverrideValue?.Val?.HasValue == true && startOverrideValue.Val.Value == 1;
                    });

                var startOverrides = new Dictionary<int, int>();
                foreach (LevelOverride levelOverride in overrides) {
                    StartOverrideNumberingValue? start = levelOverride.GetFirstChild<StartOverrideNumberingValue>();
                    if (levelOverride.LevelIndex?.HasValue == true && start?.Val?.HasValue == true && !startOverrides.ContainsKey(levelOverride.LevelIndex.Value)) {
                        startOverrides.Add(levelOverride.LevelIndex.Value, start.Val.Value);
                    }
                }

                var levels = new Dictionary<int, ListLevelDefinition>();
                if (abstractNum != null) {
                    foreach (Level level in abstractNum.Elements<Level>()) {
                        if (level.LevelIndex?.HasValue != true || levels.ContainsKey(level.LevelIndex.Value)) {
                            continue;
                        }

                        var indentation = level.GetFirstChild<PreviousParagraphProperties>()?.GetFirstChild<Indentation>();
                        NumberingSymbolRunProperties? markerProperties = level.GetFirstChild<NumberingSymbolRunProperties>();
                        var definition = new ListLevelDefinition(
                            level: level.LevelIndex.Value,
                            start: level.StartNumberingValue?.Val?.Value ?? 1,
                            numberFormat: level.NumberingFormat?.Val?.Value,
                            levelText: level.LevelText?.Val?.Value,
                            leftIndentTwips: ParseOptionalInt32(indentation?.Left?.Value),
                            hangingIndentTwips: ParseOptionalInt32(indentation?.Hanging?.Value),
                            markerFontFamily: ResolveListMarkerFontFamily(markerProperties),
                            markerBold: ReadListMarkerOnOff(markerProperties?.GetFirstChild<Bold>()),
                            markerItalic: ReadListMarkerOnOff(markerProperties?.GetFirstChild<Italic>()),
                            markerColorHex: markerProperties?.GetFirstChild<Color>()?.Val?.Value,
                            markerFontSize: ResolveListMarkerFontSize(markerProperties),
                            levelJustification: level.LevelJustification?.Val?.Value,
                            levelSuffix: level.LevelSuffix?.Val?.Value);
                        levels.Add(definition.Level, definition);
                    }
                }

                result[numberId] = new ListNumberingDefinition(
                    numberId,
                    abstractNum != null ? WordListStyles.MatchStyle(abstractNum) : WordListStyle.Custom,
                    overridesAreDefault,
                    startOverrides,
                    levels);
            }

            return result;
        }

        private static Dictionary<int, List<WordParagraph>> BuildListItemsByNumberId(WordDocument document) {
            var result = new Dictionary<int, List<WordParagraph>>();
            foreach (WordParagraph paragraph in document.EnumerateAllParagraphs()) {
                if (!paragraph.IsListItem || paragraph._listNumberId == null) {
                    continue;
                }

                int numberId = paragraph._listNumberId.Value;
                if (!result.TryGetValue(numberId, out List<WordParagraph>? items)) {
                    items = new List<WordParagraph>();
                    result[numberId] = items;
                }

                items.Add(paragraph);
            }

            return result;
        }

        private sealed class ListNumberingDefinition {
            internal ListNumberingDefinition(
                int numberId,
                WordListStyle style,
                bool overridesAreDefault,
                IReadOnlyDictionary<int, int> startOverrides,
                IReadOnlyDictionary<int, ListLevelDefinition> levels) {
                NumberId = numberId;
                Style = style;
                OverridesAreDefault = overridesAreDefault;
                StartOverrides = startOverrides;
                Levels = levels;
            }

            internal int NumberId { get; }
            internal WordListStyle Style { get; }
            internal bool OverridesAreDefault { get; }
            internal IReadOnlyDictionary<int, int> StartOverrides { get; }
            internal IReadOnlyDictionary<int, ListLevelDefinition> Levels { get; }
        }

        private readonly struct ListLevelDefinition {
            internal ListLevelDefinition(
                int level,
                int start,
                NumberFormatValues? numberFormat,
                string? levelText,
                int? leftIndentTwips,
                int? hangingIndentTwips,
                string? markerFontFamily,
                bool? markerBold,
                bool? markerItalic,
                string? markerColorHex,
                double? markerFontSize,
                LevelJustificationValues? levelJustification,
                LevelSuffixValues? levelSuffix) {
                Level = level;
                Start = start;
                NumberFormat = numberFormat;
                LevelText = levelText;
                LeftIndentTwips = leftIndentTwips;
                HangingIndentTwips = hangingIndentTwips;
                MarkerFontFamily = markerFontFamily;
                MarkerBold = markerBold;
                MarkerItalic = markerItalic;
                MarkerColorHex = markerColorHex;
                MarkerFontSize = markerFontSize;
                LevelJustification = levelJustification;
                LevelSuffix = levelSuffix;
            }

            internal int Level { get; }
            internal int Start { get; }
            internal NumberFormatValues? NumberFormat { get; }
            internal string? LevelText { get; }
            internal int? LeftIndentTwips { get; }
            internal int? HangingIndentTwips { get; }
            internal string? MarkerFontFamily { get; }
            internal bool? MarkerBold { get; }
            internal bool? MarkerItalic { get; }
            internal string? MarkerColorHex { get; }
            internal double? MarkerFontSize { get; }
            internal LevelJustificationValues? LevelJustification { get; }
            internal LevelSuffixValues? LevelSuffix { get; }
        }

        private static double? ResolveListMarkerFontSize(NumberingSymbolRunProperties? markerProperties) {
            string? value = markerProperties?.GetFirstChild<FontSize>()?.Val?.Value;
            if (!double.TryParse(value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double halfPoints) ||
                halfPoints <= 0D ||
                double.IsNaN(halfPoints) ||
                double.IsInfinity(halfPoints)) {
                return null;
            }

            return halfPoints / 2D;
        }

        private static string? ResolveListMarkerFontFamily(NumberingSymbolRunProperties? markerProperties) {
            RunFonts? runFonts = markerProperties?.GetFirstChild<RunFonts>();
            return FirstNonWhiteSpace(runFonts?.Ascii?.Value, runFonts?.HighAnsi?.Value);
        }

        private static bool? ReadListMarkerOnOff(OnOffType? value) {
            if (value == null) {
                return null;
            }

            if (value.Val == null) {
                return true;
            }

            return value.Val.Value;
        }

        private static string? FirstNonWhiteSpace(params string?[] values) {
            foreach (string? value in values) {
                if (!string.IsNullOrWhiteSpace(value)) {
                    return value;
                }
            }

            return null;
        }

        private sealed class ParagraphReferenceComparer : IEqualityComparer<WordParagraph> {
            public static readonly ParagraphReferenceComparer Instance = new();
            public bool Equals(WordParagraph? x, WordParagraph? y) => ReferenceEquals(x?._paragraph, y?._paragraph);
            public int GetHashCode(WordParagraph obj) => RuntimeHelpers.GetHashCode(obj._paragraph);
        }

        private static string BuildMarker(int level, int index, Dictionary<int, int> indices, Dictionary<int, NumberFormatValues?> formats, string? pattern) {
            if (string.IsNullOrEmpty(pattern)) {
                string formatted = FormatNumber(index, formats[level]);
                return formatted + ".";
            }

            string marker = pattern!;
            marker = marker.Replace("%CurrentLevel", FormatNumber(index, formats[level]));
            marker = Regex.Replace(marker, "%([0-9]+)", m => {
                int lvl = int.Parse(m.Groups[1].Value) - 1;
                int value = lvl == level ? index : indices.TryGetValue(lvl, out int val) ? val - 1 : 0;
                formats.TryGetValue(lvl, out NumberFormatValues? fmt);
                return FormatNumber(value, fmt);
            });
            return marker;
        }

        private static string FormatNumber(int number, NumberFormatValues? format) {
            if (format == NumberFormatValues.LowerRoman) {
                return ToRoman(number).ToLowerInvariant();
            }
            if (format == NumberFormatValues.UpperRoman) {
                return ToRoman(number);
            }
            if (format == NumberFormatValues.LowerLetter) {
                return ToAlphabeticSequence(number, uppercase: false);
            }
            if (format == NumberFormatValues.UpperLetter) {
                return ToAlphabeticSequence(number, uppercase: true);
            }
            return number.ToString();
        }

        private static string ToAlphabeticSequence(int number, bool uppercase) {
            if (number <= 0) {
                return number.ToString();
            }

            char baseCharacter = uppercase ? 'A' : 'a';
            StringBuilder sb = new();
            while (number > 0) {
                number--;
                sb.Insert(0, (char)(baseCharacter + (number % 26)));
                number /= 26;
            }

            return sb.ToString();
        }

        private static string ToRoman(int number) {
            if (number <= 0) {
                return number.ToString();
            }

            (int Value, string Symbol)[] map = new (int, string)[] {
                (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"),
                (100, "C"), (90, "XC"), (50, "L"), (40, "XL"),
                (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")
            };

            StringBuilder sb = new();
            foreach ((int value, string symbol) in map) {
                while (number >= value) {
                    sb.Append(symbol);
                    number -= value;
                }
            }

            return sb.ToString();
        }
    }
}
