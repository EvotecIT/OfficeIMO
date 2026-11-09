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
        }
    }
}
