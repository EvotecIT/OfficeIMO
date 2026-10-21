using System.Threading;

namespace OfficeIMO.Word;

public static partial class WordDocumentTraversal {
    /// <summary>
    /// Resolves a note's directly numbered paragraphs through the ordinary story-local
    /// counter owner without changing body or running-story numbering.
    /// </summary>
    internal static ListMarkerRenderScope BuildNoteListMarkersForRendering(WordDocument document,
        IReadOnlyList<WordParagraph> paragraphs, CancellationToken cancellationToken = default) {
        // Style-only numbering in note stories is not part of the qualified native
        // profile. In particular, a numbered paragraph style alone does not activate
        // the isolated note sequence observed in Word.
        IEnumerable<WordParagraph> numbered = paragraphs.Where(paragraph =>
            paragraph._paragraph.ParagraphProperties?.NumberingProperties != null);
        Dictionary<WordParagraph, ResolvedListMarker> resolved = BuildResolvedListMarkers(document,
            new[] { numbered }, cancellationToken);
        var markers = new Dictionary<WordParagraph, (int Level, string Marker)>(resolved.Comparer);
        foreach (var item in resolved) markers[item.Key] = (item.Value.Level, item.Value.Marker);
        return new ListMarkerRenderScope(markers, UseResolvedListMarkers(document, resolved, markers));
    }
}
