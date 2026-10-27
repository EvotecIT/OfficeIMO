using System.Collections.Generic;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    private static W.Endnote ResolveNativeEndnoteSeparator(WordDocument document) => ResolveNativeEndnoteStories(document).Separator;

    /// <summary>Endnotes select their distinct continuation story in both measured compatibility profiles.</summary>
    private static (W.Endnote Separator, W.Endnote ContinuationSeparator, W.Endnote? ContinuationNotice)
        ResolveNativeEndnoteStories(WordDocument document) {
        var main = document._wordprocessingDocument.MainDocumentPart;
        var registered = new HashSet<long>((main?.DocumentSettingsPart?.Settings?
            .GetFirstChild<W.EndnoteDocumentWideProperties>()?.Elements<W.EndnoteSpecialReference>() ??
            Enumerable.Empty<W.EndnoteSpecialReference>()).Where(reference => reference.Id != null)
            .Select(reference => (long)reference.Id!.Value));
        IEnumerable<W.Endnote>? source = main?.EndnotesPart?.Endnotes?.Elements<W.Endnote>();
        return (ResolveNativeLoadedNoteStory(source, registered, continuation: false),
            ResolveNativeLoadedNoteStory(source, registered, continuation: true),
            FindNativeLoadedNoteStory(source, registered, W.FootnoteEndnoteValues.ContinuationNotice));
    }

    /// <summary>Keeps endnotes in ordinary section flow while the PDF owner selects and paints accepted page fragments.</summary>
    private static bool TryRenderNativeContinuedEndnotes(INativePdfFlow flow, IReadOnlyList<PdfFootnote> endnotes,
        WordSection section, NativeNoteNumbering notes, WordPdfExportOptions? options,
        NativeDocumentDefaults defaults, NativeFontMap fonts) {
        if (endnotes.Count == 0) return true;
        var stories = ResolveNativeEndnoteStories(section._document);
        bool qualified = (section.ColumnCount ?? 1) == 1 &&
            GetNativePageFootnoteElementProfileGap(section.Elements, options, 0, 0) == null &&
            CanProjectNativePageNote(stories.Separator, section, defaults, fonts, options) &&
            CanProjectNativePageNote(stories.ContinuationSeparator, section, defaults, fonts, options) &&
            (stories.ContinuationNotice == null || CanProjectNativePageNote(stories.ContinuationNotice, section, defaults, fonts, options)) &&
            endnotes.All(note => note.Source != null && CanProjectNativePageNote(note.Source, section, defaults, fonts, options));
        if (!qualified) {
            if (options != null) AddNativeExportWarning(options, "native-endnote-continuation-approximated", "endnotes",
                "Endnote content, decorations or its body frame are outside the supported single-column continuation profile; visible content remains in section flow.");
            return false;
        }
        Action<PdfCore.PdfContentBuilder> Story(W.Endnote source) => content =>
            RenderNativeNoteStory(content, source, section, notes, options, defaults, fonts, label: null);
        flow.ContinuedFlow(content => {
            INativePdfFlow inner = new NativeSpacingCollapseFlow(new NativePdfColumnFlow(content, GetNativePageSize(section, options)));
            foreach (PdfFootnote note in endnotes) {
                options?.CancellationToken.ThrowIfCancellationRequested();
                inner.Bookmark(GetNativePageNoteDestination(GetNativeEndnoteKey(note.ReferenceId)));
                RenderNativeNoteParagraphs(inner, note.Source!, section, notes, options, defaults, fonts, note.Label);
            }
        }, Story(stories.Separator), Story(stories.ContinuationSeparator),
            stories.ContinuationNotice == null ? null : Story(stories.ContinuationNotice), reserveFinalNotice: false,
            allowContinuationSeparatorInBottomMargin: stories.ContinuationNotice == null);
        return true;
    }
}
