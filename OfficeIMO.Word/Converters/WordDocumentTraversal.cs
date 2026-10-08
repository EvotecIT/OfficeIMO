using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Runtime.CompilerServices;
using System.Threading;

namespace OfficeIMO.Word {
    /// <summary>
    /// Provides helper methods for traversing documents and resolving list markers.
    /// </summary>
    public static partial class WordDocumentTraversal {
        private static readonly AsyncLocal<ListInfoSnapshot?> ActiveListInfoSnapshot = new();

        private sealed class ListInfoSnapshot {
            internal ListInfoSnapshot(WordDocument document,
                IReadOnlyDictionary<WordParagraph, ResolvedListMarker> markers,
                IReadOnlyDictionary<WordParagraph, (int Level, string Marker)>? renderMarkers = null) {
                Document = document;
                Markers = markers;
                RenderMarkers = renderMarkers;
            }
            internal WordDocument Document { get; }
            internal IReadOnlyDictionary<WordParagraph, ResolvedListMarker> Markers { get; }
            internal IReadOnlyDictionary<WordParagraph, (int Level, string Marker)>? RenderMarkers { get; }
        }

        private sealed class ListInfoSnapshotScope : IDisposable {
            private readonly ListInfoSnapshot? _previous;
            internal ListInfoSnapshotScope(ListInfoSnapshot? previous) => _previous = previous;
            public void Dispose() => ActiveListInfoSnapshot.Value = _previous;
        }

        internal static IDisposable UseResolvedListMarkers(WordDocument document,
            IReadOnlyDictionary<WordParagraph, ResolvedListMarker> markers,
            IReadOnlyDictionary<WordParagraph, (int Level, string Marker)>? renderMarkers = null) {
            ListInfoSnapshot? previous = ActiveListInfoSnapshot.Value;
            ActiveListInfoSnapshot.Value = new ListInfoSnapshot(document, markers, renderMarkers);
            return new ListInfoSnapshotScope(previous);
        }

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

            internal ListInfo(int level, bool ordered, bool markerVisible, int start, WordNumberFormat? format, string? text, int? leftIndentTwips, int? hangingIndentTwips, string? markerFontFamily, bool? markerBold, bool? markerItalic, string? markerColorHex, double? markerFontSize, WordListLevelAlignment? levelJustification, WordListLevelSuffix? levelSuffix, int? pictureBulletId, long? markerCharacterScale = null, int? markerCharacterSpacingTwips = null) {
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
                MarkerCharacterScale = markerCharacterScale;
                MarkerCharacterSpacingTwips = markerCharacterSpacingTwips;
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
            /// <summary>Marker character width from the effective numbering level, as a percentage, when defined.</summary>
            public long? MarkerCharacterScale { get; }
            /// <summary>Marker character spacing from the effective numbering level, in twentieths of a point, when defined.</summary>
            public int? MarkerCharacterSpacingTwips { get; }
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
            internal ResolvedListMarker(ListInfo info, int index, string marker, bool useTextFont) {
                Info = info;
                Level = info.Level;
                Index = index;
                Marker = marker;
                UseTextFont = useTextFont;
                PictureBulletId = info.PictureBulletId;
            }

            internal ListInfo Info { get; }
            internal int Level { get; }
            internal int Index { get; }
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
            if (paragraph?._paragraph == null || paragraph._document == null) return null;
            ListInfoSnapshot? snapshot = ActiveListInfoSnapshot.Value;
            if (snapshot != null && ReferenceEquals(snapshot.Document, paragraph._document))
                return snapshot.Markers.TryGetValue(paragraph, out ResolvedListMarker marker) ? marker.Info : null;
            NumberingProperties? direct = paragraph._paragraph?.ParagraphProperties?.NumberingProperties;
            WordListNumberingResolver.StyleCatalog? styleCatalog = direct?.NumberingId?.Val?.Value > 0 &&
                direct.NumberingLevelReference?.Val?.Value != null
                ? null
                : WordListNumberingResolver.CreateStyleCatalog(paragraph._document);
            if (!WordListNumberingResolver.TryResolve(paragraph, out WordListNumberingResolver.ResolvedNumbering numbering, styleCatalog)) {
                return null;
            }

            Dictionary<int, ListNumberingDefinition> definitions = BuildListNumberingDefinitions(paragraph._document._wordprocessingDocument.MainDocumentPart);
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
            long? markerCharacterScale = null;
            int? markerCharacterSpacingTwips = null;
            WordListLevelAlignment? levelJustification = null;
            WordListLevelSuffix? levelSuffix = null;
            int? pictureBulletId = null;

            ListNumberingDefinition? definition = null;
            definitions.TryGetValue(numbering.NumberId, out definition);
            if (definition != null &&
                definition.StartOverrides.TryGetValue(level, out int overrideValue)) {
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
                markerCharacterScale = levelDefinition.MarkerCharacterScale;
                markerCharacterSpacingTwips = levelDefinition.MarkerCharacterSpacingTwips;
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
            return new ListInfo(level, ordered, markerVisible, start, numberFormat, levelText, leftIndentTwips, hangingIndentTwips, markerFontFamily, markerBold, markerItalic, markerColorHex, markerFontSize, levelJustification, levelSuffix, pictureBulletId, markerCharacterScale, markerCharacterSpacingTwips);
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

        internal sealed class ListMarkerRenderScope : IDisposable {
            private readonly IDisposable _listInfoScope;
            internal ListMarkerRenderScope(IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> markers,
                IDisposable listInfoScope) {
                Markers = markers;
                _listInfoScope = listInfoScope;
            }
            internal IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> Markers { get; }
            public void Dispose() => _listInfoScope.Dispose();
        }

        internal static ListMarkerRenderScope BuildListMarkersForRendering(
            WordDocument document, CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            ListInfoSnapshot? active = ActiveListInfoSnapshot.Value;
            if (ReferenceEquals(active?.Document, document) && active.RenderMarkers != null)
                return new ListMarkerRenderScope(active.RenderMarkers, new ListInfoSnapshotScope(active));
            Dictionary<WordParagraph, ResolvedListMarker> resolved = BuildResolvedListMarkers(document, cancellationToken);
            var result = new Dictionary<WordParagraph, (int, string)>(ParagraphReferenceComparer.Instance);
            foreach (KeyValuePair<WordParagraph, ResolvedListMarker> item in resolved)
                result[item.Key] = (item.Value.Level, item.Value.Marker);
            return new ListMarkerRenderScope(result, UseResolvedListMarkers(document, resolved, result));
        }

        internal static IReadOnlyList<int> GetPictureBulletFallbackIds(WordDocument document) =>
            GetPictureBulletFallbackIds(BuildResolvedListMarkers(document).Values);

        internal static IReadOnlyList<int> GetPictureBulletFallbackIds(IEnumerable<ResolvedListMarker> markers) =>
            markers.Select(marker => marker.PictureBulletId)
                .Where(id => id.HasValue).Select(id => id!.Value).Distinct().ToArray();

        internal static string ResolveTextListMarkerSuffix(WordListLevelSuffix? suffix) => suffix switch {
            WordListLevelSuffix.Nothing => string.Empty,
            WordListLevelSuffix.Space => " ",
            _ => "\t"
        };

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

        internal static Dictionary<int, ListNumberingDefinition> BuildListNumberingDefinitions(MainDocumentPart? mainPart) {
            var result = new Dictionary<int, ListNumberingDefinition>();
            Numbering? numbering = mainPart?.NumberingDefinitionsPart?.Numbering;
            if (numbering == null) {
                return result;
            }

            Dictionary<int, AbstractNum> abstracts = WordListNumberingResolver.GetCanonicalAbstractDefinitions(numbering);

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
                        ListLevelDefinition definition = CreateLevelDefinition(level.LevelIndex.Value, overrideLevel ?? level,
                            level.StartNumberingValue?.Val?.Value ?? 1, level.LevelRestart?.Val?.Value);
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
                    abstractNum?.AbstractNumberId?.Value,
                    abstractNum != null ? WordListStyles.MatchStyle(abstractNum) : WordListStyle.Custom,
                    startOverrides,
                    levels);
            }

            return result;
        }

        private static ListLevelDefinition CreateLevelDefinition(int level, Level effectiveLevel, int? abstractStart = null, int? restartLevel = null) {
            // A full w:lvlOverride replaces level formatting. Its embedded w:start does
            // not change the abstract start; w:startOverride is applied separately.
            Indentation? indentation = effectiveLevel.GetFirstChild<PreviousParagraphProperties>()?.GetFirstChild<Indentation>();
            NumberingSymbolRunProperties? markerProperties = effectiveLevel.GetFirstChild<NumberingSymbolRunProperties>();

            return new ListLevelDefinition(
                level: level,
                start: abstractStart ?? effectiveLevel.StartNumberingValue?.Val?.Value ?? 1,
                numberFormat: effectiveLevel.NumberingFormat?.Val?.Value,
                levelText: effectiveLevel.LevelText?.Val?.Value,
                leftIndentTwips: ParseOptionalInt32(indentation?.Left?.Value),
                hangingIndentTwips: ParseOptionalInt32(indentation?.Hanging?.Value),
                markerFontFamily: ResolveListMarkerFontFamily(markerProperties),
                markerBold: ReadListMarkerOnOff(markerProperties?.GetFirstChild<Bold>()),
                markerItalic: ReadListMarkerOnOff(markerProperties?.GetFirstChild<Italic>()),
                markerColorHex: markerProperties?.GetFirstChild<Color>()?.Val?.Value,
                markerFontSize: ResolveListMarkerFontSize(markerProperties),
                markerCharacterScale: markerProperties?.GetFirstChild<CharacterScale>()?.Val?.Value,
                markerCharacterSpacingTwips: markerProperties?.GetFirstChild<Spacing>()?.Val?.Value,
                levelJustification: effectiveLevel.LevelJustification?.Val?.Value,
                levelSuffix: effectiveLevel.LevelSuffix?.Val?.Value,
                pictureBulletId: effectiveLevel.GetFirstChild<LevelPictureBulletId>()?.Val?.Value,
                restartLevel: restartLevel);
        }

        private static IEnumerable<IEnumerable<WordParagraph>> EnumerateListStories(WordDocument document) {
            var seen = new HashSet<Paragraph>();

            IEnumerable<WordParagraph> EnumerateParagraphTree(Paragraph rootParagraph) {
                var pending = new Stack<Paragraph>();
                pending.Push(rootParagraph);
                while (pending.Count > 0) {
                    Paragraph paragraph = pending.Pop();
                    if (!seen.Add(paragraph)) continue;
                    yield return new WordParagraph(document, paragraph);

                    List<Paragraph> nestedParagraphs = EnumerateDirectTextBoxParagraphs(paragraph).ToList();
                    for (int index = nestedParagraphs.Count - 1; index >= 0; index--) {
                        pending.Push(nestedParagraphs[index]);
                    }
                }
            }

            IEnumerable<WordParagraph> EnumerateStory(DocumentFormat.OpenXml.OpenXmlCompositeElement? root) {
                if (root == null) yield break;
                foreach (Paragraph paragraph in root.Descendants<Paragraph>()) {
                    if (paragraph.Ancestors<TextBoxContent>().Any()) continue;
                    foreach (WordParagraph item in EnumerateParagraphTree(paragraph)) yield return item;
                }
            }

            yield return EnumerateStory(document._wordprocessingDocument.MainDocumentPart?.Document?.Body);
            foreach (WordSection section in document.Sections) {
                foreach (WordHeaderFooter? headerFooter in new WordHeaderFooter?[] { section.Header.Default, section.Header.First, section.Header.Even, section.Footer.Default, section.Footer.First, section.Footer.Even }) {
                    if (headerFooter == null) continue;
                    yield return EnumerateStory((DocumentFormat.OpenXml.OpenXmlCompositeElement?)headerFooter._header ?? headerFooter._footer);
                }
            }
        }

        private static IEnumerable<Paragraph> EnumerateDirectTextBoxParagraphs(Paragraph paragraph) {
            foreach (TextBoxContent content in EnumerateOwnedTextBoxContents(paragraph)) {
                var pending = new Stack<OpenXmlElement>();
                PushChildrenInReverse(content, pending);
                while (pending.Count > 0) {
                    OpenXmlElement element = pending.Pop();
                    if (element is TextBoxContent) {
                        continue;
                    }

                    if (element is Paragraph nestedParagraph) {
                        yield return nestedParagraph;
                        continue;
                    }

                    PushChildrenInReverse(element, pending);
                }
            }
        }

        private static IEnumerable<TextBoxContent> EnumerateOwnedTextBoxContents(Paragraph paragraph) {
            var pending = new Stack<OpenXmlElement>();
            PushChildrenInReverse(paragraph, pending);
            while (pending.Count > 0) {
                OpenXmlElement element = pending.Pop();
                if (element is TextBoxContent content) {
                    yield return content;
                    continue;
                }

                if (element is Paragraph) {
                    continue;
                }

                PushChildrenInReverse(element, pending);
            }
        }

        private static void PushChildrenInReverse(OpenXmlElement element, Stack<OpenXmlElement> pending) {
            for (OpenXmlElement? child = element.LastChild; child != null; child = child.PreviousSibling()) {
                pending.Push(child);
            }
        }

        internal sealed class ListNumberingDefinition {
            internal ListNumberingDefinition(
                int numberId,
                int? abstractNumberId,
                WordListStyle style,
                IReadOnlyDictionary<int, int> startOverrides,
                IReadOnlyDictionary<int, ListLevelDefinition> levels) {
                NumberId = numberId;
                AbstractNumberId = abstractNumberId;
                Style = style;
                StartOverrides = startOverrides;
                Levels = levels;
            }

            internal int NumberId { get; }
            internal int? AbstractNumberId { get; }
            internal WordListStyle Style { get; }
            internal IReadOnlyDictionary<int, int> StartOverrides { get; }
            internal IReadOnlyDictionary<int, ListLevelDefinition> Levels { get; }
        }

        internal readonly struct ListLevelDefinition {
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
                long? markerCharacterScale,
                int? markerCharacterSpacingTwips,
                LevelJustificationValues? levelJustification,
                LevelSuffixValues? levelSuffix,
                int? pictureBulletId,
                int? restartLevel) {
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
                MarkerCharacterScale = markerCharacterScale;
                MarkerCharacterSpacingTwips = markerCharacterSpacingTwips;
                LevelJustification = levelJustification;
                LevelSuffix = levelSuffix;
                PictureBulletId = pictureBulletId;
                RestartLevel = restartLevel;
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
            internal long? MarkerCharacterScale { get; }
            internal int? MarkerCharacterSpacingTwips { get; }
            internal LevelJustificationValues? LevelJustification { get; }
            internal LevelSuffixValues? LevelSuffix { get; }
            internal int? PictureBulletId { get; }
            // Word uses the abstract level's restart rule and ignores it in a full override.
            internal int? RestartLevel { get; }
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

    }
}
