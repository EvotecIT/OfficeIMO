using System.Threading;

namespace OfficeIMO.Word {
    public static partial class WordDocumentTraversal {
        /// <summary>
        /// Builds portable marker details for all effective list paragraphs in the document.
        /// </summary>
        internal static Dictionary<WordParagraph, ResolvedListMarker> BuildResolvedListMarkers(WordDocument document,
            CancellationToken cancellationToken = default) {
            Dictionary<WordParagraph, ResolvedListMarker> result = new(ParagraphReferenceComparer.Instance);
            VisitListCounters(document, cancellationToken, (paragraph, info, index, counters, formats) => {
                string rawMarker = !info.MarkerVisible ? string.Empty : info.PictureBulletId.HasValue ? "•"
                    : info.Ordered ? BuildMarker(info.Level, index, counters, formats, info.LevelText)
                    : info.LevelText ?? "•";
                (string marker, bool useTextFont) = NormalizeListMarker(rawMarker, info.MarkerFontFamily);
                result[paragraph] = new ResolvedListMarker(info, marker, useTextFont);
            });
            return result;
        }

        /// <summary>
        /// Builds a lookup of list numeric indices for all paragraphs in the document.
        /// The returned index is the number of the item at its nesting level,
        /// accounting for continuation and explicit restarts in each document story.
        /// </summary>
        public static Dictionary<WordParagraph, (int Level, int Index)> BuildListIndices(WordDocument document) {
            Dictionary<WordParagraph, (int, int)> result = new(ParagraphReferenceComparer.Instance);
            VisitListCounters(document, default, (paragraph, info, index, _, _) => result[paragraph] = (info.Level, index));
            return result;
        }

        private static void VisitListCounters(WordDocument document, CancellationToken cancellationToken,
            Action<WordParagraph, ListInfo, int, Dictionary<int, int>, Dictionary<int, WordNumberFormat?>> visit) {
            WordListNumberingResolver.StyleCatalog catalog = WordListNumberingResolver.CreateStyleCatalog(document);
            foreach (IEnumerable<WordParagraph> story in EnumerateListStories(document)) {
                // A numbering instance selects formatting and optional restarts. Instances
                // sharing an abstract definition otherwise advance one story-local sequence.
                var states = new Dictionary<(bool IsDefinition, int Id), ListCounterState>();
                foreach (WordParagraph paragraph in story) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (!WordListNumberingResolver.TryResolve(paragraph, out WordListNumberingResolver.ResolvedNumbering numbering,
                            catalog, cancellationToken)) continue;
                    ListInfo? resolved = GetListInfo(paragraph, numbering, catalog.ListDefinitions);
                    if (!resolved.HasValue) continue;
                    ListInfo info = resolved.Value;
                    catalog.ListDefinitions.TryGetValue(numbering.NumberId, out ListNumberingDefinition? definition);
                    var key = definition?.AbstractNumberId is int id ? (true, id) : (false, numbering.NumberId);
                    if (!states.TryGetValue(key, out ListCounterState? state)) states.Add(key, state = new ListCounterState());
                    int level = info.Level;
                    if (state.LastLevel.HasValue && level < state.LastLevel.Value) {
                        foreach (int deeper in state.Indices.Keys.Where(item => item > level).ToArray()) {
                            state.Indices.Remove(deeper);
                            state.Formats.Remove(deeper);
                        }
                    }
                    state.LastLevel = level;
                    bool firstUse = state.SeenInstanceLevels.Add((numbering.NumberId, level));
                    bool explicitRestart = firstUse && definition?.StartOverrides.ContainsKey(level) == true;
                    if (explicitRestart || !state.Indices.ContainsKey(level)) state.Indices[level] = info.Start;
                    state.Formats[level] = info.NumberFormat;
                    int current = state.Indices[level];
                    state.Indices[level] = current + 1;
                    visit(paragraph, info, current, state.Indices, state.Formats);
                }
            }
        }

        private sealed class ListCounterState {
            internal int? LastLevel { get; set; }
            internal Dictionary<int, int> Indices { get; } = new();
            internal Dictionary<int, WordNumberFormat?> Formats { get; } = new();
            internal HashSet<(int NumberId, int Level)> SeenInstanceLevels { get; } = new();
        }
    }
}
