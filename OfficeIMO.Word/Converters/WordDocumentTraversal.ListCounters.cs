using System.Threading;

namespace OfficeIMO.Word {
    public static partial class WordDocumentTraversal {
        /// <summary>
        /// Builds portable marker details for all effective list paragraphs in the document.
        /// </summary>
        internal static Dictionary<WordParagraph, ResolvedListMarker> BuildResolvedListMarkers(WordDocument document,
            CancellationToken cancellationToken = default) {
            Dictionary<WordParagraph, ResolvedListMarker> result = new(ParagraphReferenceComparer.Instance);
            int totalMarkerCharacters = 0;
            VisitListCounters(document, cancellationToken, (paragraph, info, index, counters, formats) => {
                string rawMarker = !info.MarkerVisible ? string.Empty : info.PictureBulletId.HasValue ? "•"
                    : info.Ordered ? BuildMarker(info.Level, index, counters, formats, info.LevelText)
                    : info.LevelText ?? "•";
                if (rawMarker.Length > MaximumListMarkerLength) throw ListMarkerLengthExceeded();
                (string marker, bool useTextFont) = OfficeTextListMarkerNormalizer.Normalize(rawMarker, info.MarkerFontFamily);
                if (marker.Length > MaximumDocumentListMarkerCharacters - totalMarkerCharacters)
                    throw new InvalidDataException("Word generated list markers exceed the supported document character budget.");
                totalMarkerCharacters += marker.Length;
                result[paragraph] = new ResolvedListMarker(info, index, marker, useTextFont);
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
            Action<WordParagraph, ListInfo, int, Dictionary<int, long>, Dictionary<int, WordNumberFormat?>> visit) {
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
                    int level = info.Level;
                    if (level < 0 || level >= WordListTraversal.DefaultMaximumNestingDepth)
                        throw new System.IO.InvalidDataException($"The Word list level {level} exceeds the {WordListTraversal.DefaultMaximumNestingDepth}-level traversal limit.");
                    catalog.ListDefinitions.TryGetValue(numbering.NumberId, out ListNumberingDefinition? definition);
                    var key = definition?.AbstractNumberId is int id ? (true, id) : (false, numbering.NumberId);
                    if (!states.TryGetValue(key, out ListCounterState? state)) states.Add(key, state = new ListCounterState());
                    foreach (int deeper in state.Indices.Keys.Where(item => item > level).ToArray()) {
                        int? restart = definition != null && definition.Levels.TryGetValue(deeper, out ListLevelDefinition deeperDefinition)
                            ? deeperDefinition.RestartLevel : null;
                        int trigger = restart > 0 && restart <= deeper ? restart.Value - 1 : deeper - 1;
                        if (restart != 0 && level <= trigger) {
                            state.Indices.Remove(deeper);
                        }
                    }
                    // Parent placeholders use this instance's effective formats, even
                    // when a parent was last numbered through a different instance.
                    state.Formats.Clear();
                    if (definition != null) foreach (var pair in definition.Levels)
                        state.Formats[pair.Key] = pair.Value.NumberFormat.ToOfficeEnum();
                    for (int parent = 0; parent < level; parent++) {
                        if (!state.Indices.ContainsKey(parent)) {
                            int start = definition != null && definition.Levels.TryGetValue(parent, out ListLevelDefinition parentDefinition)
                                ? parentDefinition.Start : 1;
                            // Word consumes skipped parents at their start value.
                            state.Indices[parent] = (long)start + 1;
                        }
                    }
                    bool firstUse = state.SeenInstanceLevels.Add((numbering.NumberId, level));
                    bool explicitRestart = firstUse && definition?.StartOverrides.ContainsKey(level) == true;
                    if (explicitRestart) state.Indices[level] = info.Start;
                    else if (!state.Indices.ContainsKey(level)) {
                        // After a parent advances, use the abstract start. An instance's
                        // explicit restart applies once, rather than to every later sublist.
                        state.Indices[level] = definition != null && definition.Levels.TryGetValue(level, out ListLevelDefinition levelDefinition)
                            ? levelDefinition.Start : info.Start;
                    }
                    state.Formats[level] = info.NumberFormat;
                    long next = state.Indices[level];
                    if (next > int.MaxValue || next < int.MinValue)
                        throw new System.IO.InvalidDataException("The list counter exceeds the supported 32-bit index range.");
                    int current = (int)next;
                    state.Indices[level] = next + 1;
                    visit(paragraph, info, current, state.Indices, state.Formats);
                }
            }
        }

        private sealed class ListCounterState {
            internal Dictionary<int, long> Indices { get; } = new();
            internal Dictionary<int, WordNumberFormat?> Formats { get; } = new();
            internal HashSet<(int NumberId, int Level)> SeenInstanceLevels { get; } = new();
        }
    }
}
