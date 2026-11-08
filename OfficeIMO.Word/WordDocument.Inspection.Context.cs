using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
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
            private Dictionary<Paragraph, WordComplexFieldRunVisibility.FieldState> ComplexFieldPrefixes { get; } = new();
            private HashSet<OpenXmlElement> ScannedFieldStories { get; } = new();
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

            internal WordComplexFieldRunVisibility.FieldState? ComplexFieldResultsFor(Paragraph paragraph) {
                OpenXmlElement story = paragraph.Ancestors().FirstOrDefault(element =>
                    element is Footnote or Endnote or DocumentFormat.OpenXml.Wordprocessing.Header
                        or DocumentFormat.OpenXml.Wordprocessing.Footer)
                    ?? paragraph.Ancestors().LastOrDefault() ?? paragraph;
                if (ScannedFieldStories.Add(story) && story.Descendants<FieldChar>().Any()) {
                    var visibility = new WordComplexFieldRunVisibility();
                    foreach (Paragraph candidate in story.Descendants<Paragraph>()) {
                        if (visibility.CurrentState is WordComplexFieldRunVisibility.FieldState state)
                            ComplexFieldPrefixes[candidate] = state;
                        visibility.ObserveParagraphMarkers(candidate);
                    }
                }
                return ComplexFieldPrefixes.TryGetValue(paragraph,
                    out WordComplexFieldRunVisibility.FieldState? prefix) ? prefix : null;
            }
        }
    }
}
