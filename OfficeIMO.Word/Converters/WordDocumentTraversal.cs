using DocumentFormat.OpenXml.Wordprocessing;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;

namespace OfficeIMO.Word {
    /// <summary>
    /// Provides helper methods for traversing documents and resolving list markers.
    /// </summary>
    public static class WordDocumentTraversal {
        /// <summary>
        /// Describes list information for a paragraph.
        /// </summary>
        public readonly struct ListInfo {
            /// <summary>Initializes a list descriptor using the original public constructor shape.</summary>
            public ListInfo(int level, bool ordered, int start, WordNumberFormat? format, string? text, int? leftIndentTwips, int? hangingIndentTwips)
                : this(level, ordered, markerVisible: true, start, format, text, leftIndentTwips, hangingIndentTwips, null, null, null, null, null, null, null, pictureBulletId: null) {
            }

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
            public ListInfo(int level, bool ordered, int start, WordNumberFormat? format, string? text, int? leftIndentTwips = null, int? hangingIndentTwips = null, string? markerFontFamily = null, bool? markerBold = null, bool? markerItalic = null, string? markerColorHex = null, WordListLevelAlignment? levelJustification = null, WordListLevelSuffix? levelSuffix = null)
                : this(level, ordered, markerVisible: true, start, format, text, leftIndentTwips, hangingIndentTwips, markerFontFamily, markerBold, markerItalic, markerColorHex, markerFontSize: null, levelJustification, levelSuffix, pictureBulletId: null) {
            }

            internal ListInfo(int level, bool ordered, bool markerVisible, int start, WordNumberFormat? format, string? text, int? leftIndentTwips, int? hangingIndentTwips, string? markerFontFamily, bool? markerBold, bool? markerItalic, string? markerColorHex, double? markerFontSize, WordListLevelAlignment? levelJustification, WordListLevelSuffix? levelSuffix, int? pictureBulletId) {
                Level = level;
                Ordered = ordered;
                MarkerVisible = markerVisible;
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
                PictureBulletId = pictureBulletId;
            }

            /// <summary>Zero-based nesting level.</summary>
            public int Level { get; }
            /// <summary>Indicates whether numbering is used.</summary>
            public bool Ordered { get; }
            /// <summary>Indicates whether the effective numbering level displays a marker.</summary>
            public bool MarkerVisible { get; }
            /// <summary>Starting index for the list.</summary>
            public int Start { get; }
            /// <summary>Numbering format applied to the list.</summary>
            public WordNumberFormat? NumberFormat { get; }
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
            public WordListLevelAlignment? LevelJustification { get; }
            /// <summary>Marker suffix from the numbering level, when defined.</summary>
            public WordListLevelSuffix? LevelSuffix { get; }
            /// <summary>Picture-bullet identifier from the effective numbering level, when defined.</summary>
            public int? PictureBulletId { get; }
        }

        /// <summary>
        /// Describes a list marker after Word numbering semantics and legacy symbol-font
        /// characters have been projected to portable Unicode text.
        /// </summary>
        internal readonly struct ResolvedListMarker {
            internal ResolvedListMarker(ListInfo info, string marker, bool useTextFont) {
                Info = info;
                Level = info.Level;
                Marker = marker;
                UseTextFont = useTextFont;
                PictureBulletId = info.PictureBulletId;
            }

            internal ListInfo Info { get; }
            internal int Level { get; }
            internal string Marker { get; }
            internal bool UseTextFont { get; }
            internal int? PictureBulletId { get; }
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
            if (paragraph == null ||
                !WordListNumberingResolver.TryResolve(paragraph, out WordListNumberingResolver.ResolvedNumbering numbering,
                    WordListNumberingResolver.CreateStyleCatalog(paragraph._document))) {
                return null;
            }

            Dictionary<int, ListNumberingDefinition> definitions = BuildListNumberingDefinitions(paragraph._document);
            return GetListInfo(paragraph, numbering, definitions);
        }

        private static ListInfo? GetListInfo(WordParagraph paragraph, IReadOnlyDictionary<int, ListNumberingDefinition> definitions, WordListNumberingResolver.StyleCatalog styleCatalog) {
            if (paragraph == null ||
                !WordListNumberingResolver.TryResolve(paragraph, out WordListNumberingResolver.ResolvedNumbering numbering, styleCatalog)) {
                return null;
            }

            return GetListInfo(paragraph, numbering, definitions);
        }

        private static ListInfo? GetListInfo(
            WordParagraph paragraph,
            WordListNumberingResolver.ResolvedNumbering numbering,
            IReadOnlyDictionary<int, ListNumberingDefinition> definitions) {
            int level = numbering.Level;
            int? overrideStart = null;
            int start = 1;
            WordNumberFormat? numberFormat = null;
            string? levelText = null;
            int? leftIndentTwips = null;
            int? hangingIndentTwips = null;
            string? markerFontFamily = null;
            bool? markerBold = null;
            bool? markerItalic = null;
            string? markerColorHex = null;
            double? markerFontSize = null;
            WordListLevelAlignment? levelJustification = null;
            WordListLevelSuffix? levelSuffix = null;
            int? pictureBulletId = null;

            ListNumberingDefinition? definition = null;
            definitions.TryGetValue(numbering.NumberId, out definition);
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
                numberFormat = levelDefinition.NumberFormat.ToOfficeEnum();
                levelText = levelDefinition.LevelText;
                leftIndentTwips = levelDefinition.LeftIndentTwips;
                hangingIndentTwips = levelDefinition.HangingIndentTwips;
                markerFontFamily = levelDefinition.MarkerFontFamily;
                markerBold = levelDefinition.MarkerBold;
                markerItalic = levelDefinition.MarkerItalic;
                markerColorHex = levelDefinition.MarkerColorHex;
                markerFontSize = levelDefinition.MarkerFontSize;
                levelJustification = levelDefinition.LevelJustification.ToOfficeEnum();
                levelSuffix = levelDefinition.LevelSuffix.ToOfficeEnum();
                pictureBulletId = levelDefinition.PictureBulletId;
            }

            bool markerVisible = pictureBulletId.HasValue || numberFormat != WordNumberFormat.None;
            bool ordered = pictureBulletId.HasValue ? false : numberFormat.HasValue
                ? numberFormat.Value != WordNumberFormat.Bullet && numberFormat.Value != WordNumberFormat.None
                : definition?.Style switch {
                    WordListStyle.Bulleted => false,
                    WordListStyle.BulletedChars => false,
                    _ => true,
                };
            return new ListInfo(level, ordered, markerVisible, start, numberFormat, levelText, leftIndentTwips, hangingIndentTwips, markerFontFamily, markerBold, markerItalic, markerColorHex, markerFontSize, levelJustification, levelSuffix, pictureBulletId);
        }

        private static int? ParseOptionalInt32(string? value) {
            return int.TryParse(value, out int parsed) ? parsed : null;
        }

        /// <summary>
        /// Builds a lookup of list markers for all paragraphs in the document.
        /// </summary>
        public static Dictionary<WordParagraph, (int Level, string Marker)> BuildListMarkers(WordDocument document) {
            Dictionary<WordParagraph, ResolvedListMarker> resolved = BuildResolvedListMarkers(document);
            var result = new Dictionary<WordParagraph, (int, string)>(ParagraphReferenceComparer.Instance);
            foreach (KeyValuePair<WordParagraph, ResolvedListMarker> item in resolved) {
                result[item.Key] = (item.Value.Level, item.Value.Marker);
            }

            return result;
        }

        /// <summary>
        /// Builds portable marker details for all effective list paragraphs in the document.
        /// </summary>
        internal static Dictionary<WordParagraph, ResolvedListMarker> BuildResolvedListMarkers(WordDocument document) {
            Dictionary<WordParagraph, ResolvedListMarker> result = new(ParagraphReferenceComparer.Instance);
            WordListNumberingResolver.StyleCatalog styleCatalog = WordListNumberingResolver.CreateStyleCatalog(document);
            Dictionary<int, ListNumberingDefinition> definitions = BuildListNumberingDefinitions(document);
            Dictionary<int, List<WordParagraph>> itemsByNumberId = BuildListItemsByNumberId(document, styleCatalog);

            foreach (KeyValuePair<int, List<WordParagraph>> listItems in itemsByNumberId) {
                Dictionary<int, int> indices = new();
                Dictionary<int, WordNumberFormat?> formats = new();
                int lastLevel = 0;
                bool first = true;
                foreach (WordParagraph item in listItems.Value) {
                    ListInfo? info = GetListInfo(item, definitions, styleCatalog);
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

                    string rawMarker;
                    if (!info.Value.MarkerVisible) {
                        rawMarker = string.Empty;
                    } else if (info.Value.PictureBulletId.HasValue) {
                        rawMarker = "•";
                    } else {
                        rawMarker = info.Value.Ordered
                            ? BuildMarker(level, currentIndex, indices, formats, info.Value.LevelText)
                            : (info.Value.LevelText ?? "•");
                    }

                    (string marker, bool useTextFont) = NormalizeListMarker(rawMarker, info.Value.MarkerFontFamily);
                    result[item] = new ResolvedListMarker(info.Value, marker, useTextFont);
                }
            }

            return result;
        }

        private static (string Marker, bool UseTextFont) NormalizeListMarker(string marker, string? fontFamily) {
            string trimmed = marker.Trim();
            if (trimmed.Length == 0) {
                return (string.Empty, false);
            }

            if (string.Equals(fontFamily, "Symbol", StringComparison.OrdinalIgnoreCase)) {
                return trimmed switch {
                    "\uf0b7" => ("•", true),
                    "\u00b7" => ("•", true),
                    _ => (marker, false)
                };
            }

            if (string.Equals(fontFamily, "Wingdings", StringComparison.OrdinalIgnoreCase)) {
                return trimmed switch {
                    "\uf0a7" => ("▪", true),
                    _ => (marker, false)
                };
            }

            return (marker, false);
        }

        internal static bool ShouldUseTextFontForMarker(ListInfo? info, string marker) {
            if (info == null) return false;
            if (info.Value.PictureBulletId.HasValue) return true;
            (string normalized, bool useTextFont) = NormalizeListMarker(info.Value.LevelText ?? string.Empty, info.Value.MarkerFontFamily);
            return useTextFont && string.Equals(normalized, marker, StringComparison.Ordinal);
        }

        /// <summary>
        /// Builds a lookup of list numeric indices for all paragraphs in the document.
        /// The returned index is the 1-based number of the item at its nesting level,
        /// accounting for list continuation across unrelated content.
        /// </summary>
        public static Dictionary<WordParagraph, (int Level, int Index)> BuildListIndices(WordDocument document) {
            Dictionary<WordParagraph, (int, int)> result = new(ParagraphReferenceComparer.Instance);
            WordListNumberingResolver.StyleCatalog styleCatalog = WordListNumberingResolver.CreateStyleCatalog(document);
            Dictionary<int, ListNumberingDefinition> definitions = BuildListNumberingDefinitions(document);
            Dictionary<int, List<WordParagraph>> itemsByNumberId = BuildListItemsByNumberId(document, styleCatalog);

            foreach (KeyValuePair<int, List<WordParagraph>> listItems in itemsByNumberId) {
                // Track current numbering per level within this list
                Dictionary<int, int> indices = new();
                int lastLevel = 0;
                bool first = true;
                foreach (WordParagraph item in listItems.Value) {
                    ListInfo? info = GetListInfo(item, definitions, styleCatalog);
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
                var levelOverrides = new Dictionary<int, Level>();
                foreach (LevelOverride levelOverride in overrides) {
                    StartOverrideNumberingValue? start = levelOverride.GetFirstChild<StartOverrideNumberingValue>();
                    if (levelOverride.LevelIndex?.HasValue == true && start?.Val?.HasValue == true && !startOverrides.ContainsKey(levelOverride.LevelIndex.Value)) {
                        startOverrides.Add(levelOverride.LevelIndex.Value, start.Val.Value);
                    }
                    Level? overrideLevel = levelOverride.GetFirstChild<Level>();
                    if (levelOverride.LevelIndex?.HasValue == true && overrideLevel != null && !levelOverrides.ContainsKey(levelOverride.LevelIndex.Value)) {
                        levelOverrides.Add(levelOverride.LevelIndex.Value, overrideLevel);
                    }
                }

                var levels = new Dictionary<int, ListLevelDefinition>();
                if (abstractNum != null) {
                    foreach (Level level in abstractNum.Elements<Level>()) {
                        if (level.LevelIndex?.HasValue != true || levels.ContainsKey(level.LevelIndex.Value)) {
                            continue;
                        }

                        levelOverrides.TryGetValue(level.LevelIndex.Value, out Level? overrideLevel);
                        ListLevelDefinition definition = CreateLevelDefinition(level.LevelIndex.Value, overrideLevel ?? level);
                        levels.Add(definition.Level, definition);
                    }
                }
                foreach (KeyValuePair<int, Level> levelOverride in levelOverrides) {
                    if (!levels.ContainsKey(levelOverride.Key)) {
                        levels.Add(levelOverride.Key, CreateLevelDefinition(levelOverride.Key, levelOverride.Value));
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

        private static ListLevelDefinition CreateLevelDefinition(int level, Level effectiveLevel) {
            // A full w:lvlOverride replaces the abstract level. A start-only override is
            // applied separately, leaving this definition on the abstract level.
            Indentation? indentation = effectiveLevel.GetFirstChild<PreviousParagraphProperties>()?.GetFirstChild<Indentation>();
            NumberingSymbolRunProperties? markerProperties = effectiveLevel.GetFirstChild<NumberingSymbolRunProperties>();

            return new ListLevelDefinition(
                level: level,
                start: effectiveLevel.StartNumberingValue?.Val?.Value ?? 1,
                numberFormat: effectiveLevel.NumberingFormat?.Val?.Value,
                levelText: effectiveLevel.LevelText?.Val?.Value,
                leftIndentTwips: ParseOptionalInt32(indentation?.Left?.Value),
                hangingIndentTwips: ParseOptionalInt32(indentation?.Hanging?.Value),
                markerFontFamily: ResolveListMarkerFontFamily(markerProperties),
                markerBold: ReadListMarkerOnOff(markerProperties?.GetFirstChild<Bold>()),
                markerItalic: ReadListMarkerOnOff(markerProperties?.GetFirstChild<Italic>()),
                markerColorHex: markerProperties?.GetFirstChild<Color>()?.Val?.Value,
                markerFontSize: ResolveListMarkerFontSize(markerProperties),
                levelJustification: effectiveLevel.LevelJustification?.Val?.Value,
                levelSuffix: effectiveLevel.LevelSuffix?.Val?.Value,
                pictureBulletId: effectiveLevel.GetFirstChild<LevelPictureBulletId>()?.Val?.Value);
        }

        private static Dictionary<int, List<WordParagraph>> BuildListItemsByNumberId(WordDocument document, WordListNumberingResolver.StyleCatalog styleCatalog) {
            var result = new Dictionary<int, List<WordParagraph>>();
            foreach (WordParagraph paragraph in EnumerateListParagraphs(document)) {
                if (!WordListNumberingResolver.TryResolve(paragraph, out WordListNumberingResolver.ResolvedNumbering numbering, styleCatalog)) {
                    continue;
                }

                int numberId = numbering.NumberId;
                if (!result.TryGetValue(numberId, out List<WordParagraph>? items)) {
                    items = new List<WordParagraph>();
                    result[numberId] = items;
                }

                items.Add(paragraph);
            }

            return result;
        }

        private static IEnumerable<WordParagraph> EnumerateListParagraphs(WordDocument document) {
            var seen = new HashSet<Paragraph>();
            var boxesByAnchor = new Dictionary<Paragraph, List<WordTextBox>>();
            void IndexTextBoxes(IEnumerable<WordTextBox> textBoxes) {
                foreach (WordTextBox textBox in textBoxes) {
                    if (textBox.AnchorParagraph is not Paragraph anchor) continue;
                    if (!boxesByAnchor.TryGetValue(anchor, out List<WordTextBox>? boxes)) {
                        boxes = new List<WordTextBox>();
                        boxesByAnchor.Add(anchor, boxes);
                    }
                    boxes.Add(textBox);
                }
            }

            IndexTextBoxes(document.TextBoxes);
            foreach (WordSection section in document.Sections) {
                foreach (WordHeaderFooter? headerFooter in new WordHeaderFooter?[] { section.Header.Default, section.Header.First, section.Header.Even, section.Footer.Default, section.Footer.First, section.Footer.Even }) {
                    if (headerFooter != null) IndexTextBoxes(headerFooter.TextBoxes);
                }
            }

            IEnumerable<WordParagraph> EnumerateStory(DocumentFormat.OpenXml.OpenXmlCompositeElement? root) {
                if (root == null) yield break;
                foreach (Paragraph paragraph in root.Descendants<Paragraph>()) {
                    if (paragraph.Ancestors<TextBoxContent>().Any() || !seen.Add(paragraph)) continue;
                    yield return new WordParagraph(document, paragraph);
                    if (!boxesByAnchor.TryGetValue(paragraph, out List<WordTextBox>? boxes)) continue;
                    foreach (WordTextBox box in boxes) {
                        foreach (WordParagraph inner in box.Paragraphs) {
                            if (inner._paragraph != null && seen.Add(inner._paragraph)) yield return inner;
                        }
                    }
                }
            }

            foreach (WordParagraph paragraph in EnumerateStory(document._wordprocessingDocument.MainDocumentPart?.Document?.Body)) {
                yield return paragraph;
            }
            foreach (WordSection section in document.Sections) {
                foreach (WordHeaderFooter? headerFooter in new WordHeaderFooter?[] { section.Header.Default, section.Header.First, section.Header.Even, section.Footer.Default, section.Footer.First, section.Footer.Even }) {
                    if (headerFooter == null) continue;
                    foreach (WordParagraph paragraph in EnumerateStory((DocumentFormat.OpenXml.OpenXmlCompositeElement?)headerFooter._header ?? headerFooter._footer)) {
                        yield return paragraph;
                    }
                }
            }
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
                LevelSuffixValues? levelSuffix,
                int? pictureBulletId) {
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
                PictureBulletId = pictureBulletId;
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
            internal int? PictureBulletId { get; }
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

        private static string BuildMarker(int level, int index, Dictionary<int, int> indices, Dictionary<int, WordNumberFormat?> formats, string? pattern) {
            if (string.IsNullOrEmpty(pattern)) {
                string formatted = FormatNumber(index, formats[level]);
                return formatted + ".";
            }

            string marker = pattern!;
            marker = marker.Replace("%CurrentLevel", FormatNumber(index, formats[level]));
            marker = Regex.Replace(marker, "%([0-9]+)", m => {
                if (!int.TryParse(m.Groups[1].Value, out int placeholderLevel) || placeholderLevel <= 0) {
                    return m.Value;
                }
                int lvl = placeholderLevel - 1;
                int value = lvl == level ? index : indices.TryGetValue(lvl, out int val) ? val - 1 : 0;
                formats.TryGetValue(lvl, out WordNumberFormat? fmt);
                return FormatNumber(value, fmt);
            });
            return marker;
        }

        private static string FormatNumber(int number, WordNumberFormat? format) {
            if (format == WordNumberFormat.LowerRoman) {
                return ToRoman(number).ToLowerInvariant();
            }
            if (format == WordNumberFormat.UpperRoman) {
                return ToRoman(number);
            }
            if (format == WordNumberFormat.LowerLetter) {
                return ToAlphabeticSequence(number, uppercase: false);
            }
            if (format == WordNumberFormat.UpperLetter) {
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
