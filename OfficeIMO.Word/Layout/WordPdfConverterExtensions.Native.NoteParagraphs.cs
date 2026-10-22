using System.Collections.Generic;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    /// <summary>
    /// Projects a note through the same paragraph, decoration, inline-object and
    /// hidden-boundary owners used by ordinary Word flow.
    /// </summary>
    private static void RenderNativeNoteParagraphs(INativePdfFlow flow, W.FootnoteEndnoteType source,
        WordSection section, NativeNoteNumbering notes, WordToPdfOptions? options,
        NativeDocumentDefaults defaults, NativeFontMap fonts, string? label) {
        defaults = CreateNativeNoteStoryDefaults(defaults, section, label);
        IReadOnlyList<WordParagraph> paragraphs = source.Elements<W.Paragraph>()
            .Select(paragraph => new WordParagraph(section._document, paragraph)).ToArray();
        IReadOnlyList<WordElement> elements = paragraphs.Cast<WordElement>().ToArray();
        using WordDocumentTraversal.ListMarkerRenderScope markers = WordDocumentTraversal
            .BuildNoteListMarkersForRendering(section._document, paragraphs, options?.CancellationToken ?? default);
        (int Level, string Marker)? GetMarker(WordParagraph paragraph) =>
            markers.Markers.TryGetValue(paragraph, out var marker) ? marker : null;
        var destinations = new Dictionary<W.Paragraph, string>();
        for (int index = 0; index < elements.Count; index++) {
            options?.CancellationToken.ThrowIfCancellationRequested();
            if (TryRenderNativeJoinedParagraphs(flow, elements, ref index, GetMarker,
                notes, options, defaults, fonts)) continue;
            RenderNativeElement(flow, elements[index], section, GetMarker, Array.Empty<int>(), notes, options,
                Array.Empty<NativeTableOfContentsEntry>(), destinations, contentWidth: null, defaults, fonts,
                nextElement: GetNextNativeRenderableElement(elements, index));
        }
    }
}
