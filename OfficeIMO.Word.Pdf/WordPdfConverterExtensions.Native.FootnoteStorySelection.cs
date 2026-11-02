using DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    /// <summary>Resolves the loaded stories actually used by this document's page-note layout.</summary>
    private static (W.Footnote Separator, W.Footnote ContinuationSeparator, W.Footnote? ContinuationNotice) ResolveNativePageFootnoteStories(
        WordDocument document) {
        var main = document._wordprocessingDocument.MainDocumentPart;
        W.Footnotes? source = main?.FootnotesPart?.Footnotes;
        var registered = new HashSet<long>((main?.DocumentSettingsPart?.Settings?
            .GetFirstChild<W.FootnoteDocumentWideProperties>()?.Elements<W.FootnoteSpecialReference>() ??
            Enumerable.Empty<W.FootnoteSpecialReference>()).Where(reference => reference.Id != null)
            .Select(reference => (long)reference.Id!.Value));
        W.Footnote separator = ResolveNativeLoadedFootnoteStory(source, registered, continuation: false);
        // Native modern layout reuses the ordinary separator on continued pages.
        // An explicitly loaded empty story remains empty; only an absent loaded
        // story receives the detached implicit default.
        return (separator, UsesModernNativeWordLayout(document) ? separator :
            ResolveNativeLoadedFootnoteStory(source, registered, continuation: true),
            FindNativeLoadedFootnoteStory(source, registered, W.FootnoteEndnoteValues.ContinuationNotice));
    }

    private static W.Footnote ResolveNativeLoadedFootnoteStory(W.Footnotes? source,
        ISet<long> registered, bool continuation) => ResolveNativeLoadedNoteStory(
            source?.Elements<W.Footnote>(), registered, continuation);

    private static TNote ResolveNativeLoadedNoteStory<TNote>(IEnumerable<TNote>? source,
        ISet<long> registered, bool continuation) where TNote : W.FootnoteEndnoteType, new() {
        W.FootnoteEndnoteValues type = continuation
            ? W.FootnoteEndnoteValues.ContinuationSeparator : W.FootnoteEndnoteValues.Separator;
        TNote? stored = FindNativeLoadedNoteStory(source, registered, type);
        if (stored != null) return stored;
        // Registration decides which stored definitions Word loads. Unregistered
        // definitions must not supply content, font dependencies or admission
        // restrictions. Construct the existing default without editing the part.
        OpenXmlElement mark = continuation ? new W.ContinuationSeparatorMark() : new W.SeparatorMark();
        var implicitStory = new TNote { Type = type, Id = continuation ? 0 : -1 };
        implicitStory.Append(new W.Paragraph(
            new W.ParagraphProperties(new W.SpacingBetweenLines { Line = "240", LineRule = W.LineSpacingRuleValues.Auto }),
            new W.Run(mark)));
        return implicitStory;
    }

    // Notices are optional in both compatibility profiles. A missing or
    // unregistered notice remains absent instead of acquiring an implicit story.
    private static W.Footnote? FindNativeLoadedFootnoteStory(W.Footnotes? source,
        ISet<long> registered, W.FootnoteEndnoteValues type) =>
        FindNativeLoadedNoteStory(source?.Elements<W.Footnote>(), registered, type);

    private static TNote? FindNativeLoadedNoteStory<TNote>(IEnumerable<TNote>? source,
        ISet<long> registered, W.FootnoteEndnoteValues type) where TNote : W.FootnoteEndnoteType =>
        source?.FirstOrDefault(note => note.Type?.Value == type &&
                note.Id?.Value is long id && registered.Contains(id));
}
