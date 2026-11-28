using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    private static OfficeXmlNode ResolveNativeEndnoteSeparator(WordDocument document) => ResolveNativeEndnoteStories(document).Separator;

    /// <summary>Endnotes select their distinct continuation story in both measured compatibility profiles.</summary>
    private static (OfficeXmlNode Separator, OfficeXmlNode ContinuationSeparator, OfficeXmlNode ContinuationNotice)
        ResolveNativeEndnoteStories(WordDocument document) {
        ISet<long> registered = GetNativeRegisteredNoteStories(document, "endnotePr", "endnote");
        IEnumerable<OfficeXmlNode> source = document.EndnotesNode().WElements("endnote");
        return (ResolveNativeLoadedNoteStory(source, registered, continuation: false, endnote: true),
            ResolveNativeLoadedNoteStory(source, registered, continuation: true, endnote: true),
            FindNativeLoadedNoteStory(source, registered, "continuationNotice"));
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
            (stories.ContinuationNotice.IsNone || CanProjectNativePageNote(stories.ContinuationNotice, section, defaults, fonts, options)) &&
            endnotes.All(note => note.Source.Exists && CanProjectNativePageNote(note.Source, section, defaults, fonts, options));
        if (!qualified) {
            if (options != null) AddNativeExportWarning(options, "native-endnote-continuation-approximated", "endnotes",
                "Endnote content, decorations or its body frame are outside the supported single-column continuation profile; visible content remains in section flow.");
            return false;
        }
        Action<PdfCore.PdfContentBuilder> Story(OfficeXmlNode source) => content =>
            RenderNativeNoteStory(content, source, section, notes, options, defaults, fonts, label: null);
        flow.ContinuedFlow(content => {
            INativePdfFlow inner = new NativeSpacingCollapseFlow(new NativePdfColumnFlow(content, GetNativePageSize(section, options)));
            foreach (PdfFootnote note in endnotes) {
                options?.CancellationToken.ThrowIfCancellationRequested();
                inner.Bookmark(GetNativePageNoteDestination(GetNativeEndnoteKey(note.ReferenceId)));
                RenderNativeNoteParagraphs(inner, note.Source, section, notes, options, defaults, fonts, note.Label);
            }
        }, Story(stories.Separator), Story(stories.ContinuationSeparator),
            stories.ContinuationNotice.IsNone ? null : Story(stories.ContinuationNotice), reserveFinalNotice: false,
            allowContinuationSeparatorInBottomMargin: stories.ContinuationNotice.IsNone);
        return true;
    }
}
