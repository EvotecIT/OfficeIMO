using System.Collections.Generic;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    /// <summary>
    /// Projects supported note blocks in source order through the same paragraph,
    /// table, inline-object and hidden-boundary owners used by ordinary Word flow.
    /// </summary>
    private static void RenderNativeNoteParagraphs(INativePdfFlow flow, W.FootnoteEndnoteType source,
        WordSection section, NativeNoteNumbering notes, WordToPdfOptions? options,
        NativeDocumentDefaults defaults, NativeFontMap fonts, string? label) {
        defaults = CreateNativeNoteStoryDefaults(defaults, section, label);
        IReadOnlyList<WordElement> elements = GetNativeNoteElements(section._document, source);
        IReadOnlyList<WordParagraph> paragraphs = EnumerateNativeNoteParagraphs(elements).ToArray();
        using WordDocumentTraversal.ListMarkerRenderScope markers = WordDocumentTraversal
            .BuildNoteListMarkersForRendering(section._document, paragraphs, options?.CancellationToken ?? default);
        (int Level, string Marker)? GetMarker(WordParagraph paragraph) =>
            markers.Markers.TryGetValue(paragraph, out var marker) ? marker : null;
        var destinations = new Dictionary<W.Paragraph, string>();
        double? contentWidth = (section.ColumnCount ?? 1) == 1
            ? GetNativeFootnoteSectionWidth(section, options) : null;
        int rejectedJoinUntil = -1;
        for (int index = 0; index < elements.Count; index++) {
            options?.CancellationToken.ThrowIfCancellationRequested();
            if (TryRenderNativeJoinedParagraphs(flow, elements, ref index, ref rejectedJoinUntil, GetMarker,
                notes, options, defaults, fonts)) continue;
            RenderNativeElement(flow, elements[index], section, GetMarker, Array.Empty<int>(), notes, options,
                Array.Empty<NativeTableOfContentsEntry>(), destinations, contentWidth, defaults, fonts,
                nextElement: GetNextNativeRenderableElement(elements, index));
        }
    }
}
