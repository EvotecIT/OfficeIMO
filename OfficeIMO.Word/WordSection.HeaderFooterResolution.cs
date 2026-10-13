using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

public partial class WordSection {
    internal WordHeader? ResolveDefaultHeader() => ResolveHeader(HeaderFooterValues.Default);
    internal WordHeader? ResolveFirstHeader() => ResolveHeader(HeaderFooterValues.First);
    internal WordHeader? ResolveEvenHeader() => ResolveHeader(HeaderFooterValues.Even);
    internal WordFooter? ResolveDefaultFooter() => ResolveFooter(HeaderFooterValues.Default);
    internal WordFooter? ResolveFirstFooter() => ResolveFooter(HeaderFooterValues.First);
    internal WordFooter? ResolveEvenFooter() => ResolveFooter(HeaderFooterValues.Even);

    // Omitted section-local references inherit the preceding story of that
    // role. An explicit reference, including an empty part, ends the search.
    // Keep the authored slots unchanged so callers can still inspect linkage.
    private WordHeader? ResolveHeader(HeaderFooterValues kind) {
        for (int index = GetSectionOrdinal(); index >= 0; index--) {
            WordSection section = _document.Sections[index];
            if (!WordHeadersAndFooters.GetHeaderReference(_document, section, kind)) continue;
            return kind == HeaderFooterValues.First ? section.Header?.First
                : kind == HeaderFooterValues.Even ? section.Header?.Even : section.Header?.Default;
        }
        return null;
    }

    private WordFooter? ResolveFooter(HeaderFooterValues kind) {
        for (int index = GetSectionOrdinal(); index >= 0; index--) {
            WordSection section = _document.Sections[index];
            if (!WordHeadersAndFooters.GetFooterReference(_document, section, kind)) continue;
            return kind == HeaderFooterValues.First ? section.Footer?.First
                : kind == HeaderFooterValues.Even ? section.Footer?.Even : section.Footer?.Default;
        }
        return null;
    }
}
