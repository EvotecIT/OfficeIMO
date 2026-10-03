using System.Threading;

namespace OfficeIMO.Epub;

internal static partial class EpubReader {
    // Repeated spine positions share parsing results, including failures. Cache entries are
    // released at their last position; raw markup is cached only within the retention budget.
    private static ParsedChapterEntry ParseChapterEntry(ChapterCandidate candidate, EpubReadOptions options,
        bool canRetainHtml, CancellationToken token) {
        string html;
        try {
            html = ReadEntryText(candidate.Entry, options.MaxChapterBytes, token);
        } catch (DecoderFallbackException) {
            return new ParsedChapterEntry(ChapterMarkupInfo.Empty, null, "epub.chapter.invalid-encoding",
                $"Skipped chapter '{candidate.Path}' because its character encoding is invalid.");
        }
        if (!TryReadChapterMarkup(html, out ChapterMarkupInfo markup, token))
            return new ParsedChapterEntry(ChapterMarkupInfo.Empty, null, "epub.chapter.invalid-xhtml",
                $"Skipped chapter '{candidate.Path}' because chapter markup is not valid XML/XHTML.");
        return new ParsedChapterEntry(markup, options.IncludeRawHtml && canRetainHtml ? html : null, null, null);
    }

    private sealed record ParsedChapterEntry(ChapterMarkupInfo Markup, string? Html, string? ErrorCode, string? ErrorMessage);
}
