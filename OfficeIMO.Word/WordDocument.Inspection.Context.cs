using System.Threading;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        private sealed class InspectionExpansionContext {
            internal InspectionExpansionContext(IReadOnlyDictionary<string, string?> paragraphStyleNames,
                IReadOnlyDictionary<WordParagraph, WordDocumentTraversal.ResolvedListMarker> listMarkers) {
                ParagraphStyleNames = paragraphStyleNames;
                ListMarkers = listMarkers;
            }

            internal HashSet<string> ActiveNoteKeys { get; } = new(StringComparer.Ordinal);
            internal bool IncludeImageBytes { get; set; } = true;
            internal CancellationToken CancellationToken { get; set; }
            internal IReadOnlyDictionary<string, string?> ParagraphStyleNames { get; }
            internal IReadOnlyDictionary<WordParagraph, WordDocumentTraversal.ResolvedListMarker> ListMarkers { get; }
            private Dictionary<(string? CharacterStyleId, string? ParagraphStyleId, bool? Caps, bool? SmallCaps), WordCapsStyle> CapsStyles { get; } = new();

            /// <summary>Reuses resolved caps only within this read-only inspection; later inspections observe source edits.</summary>
            internal WordCapsStyle CapsStyleFor(WordParagraph paragraph) {
                var direct = paragraph.GetDirectCapsOverrides();
                if (direct.Caps.HasValue && direct.SmallCaps.HasValue)
                    return paragraph.ResolveEffectiveCapsStyle(direct);
                var key = (CharacterStyleId: paragraph.CharacterStyleId, ParagraphStyleId: paragraph.StyleId,
                    Caps: direct.Caps, SmallCaps: direct.SmallCaps);
                if (!CapsStyles.TryGetValue(key, out WordCapsStyle style)) {
                    style = paragraph.ResolveEffectiveCapsStyle(direct);
                    CapsStyles.Add(key, style);
                }
                return style;
            }

            private Dictionary<int, WordList>? _firstListByNumberId;
            private readonly Dictionary<int, string?> _listStyleNames = new();

            /// <summary>
            /// Same result as <see cref="WordParagraph.ListStyle"/>, which builds every list in the document for each
            /// call. The document is read-only for the inspection, so its lists are built once and the first list
            /// with each number id answers every paragraph that uses it.
            /// </summary>
            internal string? ListStyleNameFor(WordParagraph paragraph) {
                if (!paragraph.IsListItem) return null;
                int? numberId = paragraph._listNumberId;
                if (numberId == null || paragraph._document == null) return null;
                if (_firstListByNumberId == null) {
                    _firstListByNumberId = new Dictionary<int, WordList>();
                    foreach (WordList list in paragraph._document.Lists) {
                        if (!_firstListByNumberId.ContainsKey(list._numberId)) _firstListByNumberId.Add(list._numberId, list);
                    }
                }
                if (!_listStyleNames.TryGetValue(numberId.Value, out string? name)) {
                    name = _firstListByNumberId.TryGetValue(numberId.Value, out WordList? first) ? first.Style.ToString() : null;
                    _listStyleNames.Add(numberId.Value, name);
                }
                return name;
            }
        }
    }
}
