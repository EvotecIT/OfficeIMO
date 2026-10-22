using System.Collections.Generic;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    private static W.Endnote ResolveNativeEndnoteSeparator(WordDocument document) {
        var main = document._wordprocessingDocument.MainDocumentPart;
        var registered = new HashSet<long>((main?.DocumentSettingsPart?.Settings?
            .GetFirstChild<W.EndnoteDocumentWideProperties>()?.Elements<W.EndnoteSpecialReference>() ??
            Enumerable.Empty<W.EndnoteSpecialReference>()).Where(reference => reference.Id != null)
            .Select(reference => (long)reference.Id!.Value));
        // Endnotes use the same registration/explicit-empty/implicit-story
        // selection owner as footnotes. Their continuation stories remain a
        // separate qualification boundary of ordinary section flow.
        return ResolveNativeLoadedNoteStory(main?.EndnotesPart?.Endnotes?.Elements<W.Endnote>(),
            registered, continuation: false);
    }
}
