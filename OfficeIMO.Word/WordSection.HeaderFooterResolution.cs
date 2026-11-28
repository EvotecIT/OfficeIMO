namespace OfficeIMO.Word;

public partial class WordSection {
    internal WordHeader? ResolveDefaultHeader() => ResolveHeader(WordHeaderFooterType.Default);
    internal WordHeader? ResolveFirstHeader() => ResolveHeader(WordHeaderFooterType.First);
    internal WordHeader? ResolveEvenHeader() => ResolveHeader(WordHeaderFooterType.Even);
    internal WordFooter? ResolveDefaultFooter() => ResolveFooter(WordHeaderFooterType.Default);
    internal WordFooter? ResolveFirstFooter() => ResolveFooter(WordHeaderFooterType.First);
    internal WordFooter? ResolveEvenFooter() => ResolveFooter(WordHeaderFooterType.Even);

    // Omitted section-local references inherit the preceding story of that
    // role. An explicit reference, including an empty part, ends the search.
    // Keep the authored slots unchanged so callers can still inspect linkage.
    private WordHeader? ResolveHeader(WordHeaderFooterType kind) {
        var snapshot = WordHeaderFooterResolutionSnapshot.Current(_document);
        if (snapshot != null && snapshot.TryGet(this, out var stories)) return stories.Header(kind);
        for (int index = GetSectionOrdinal(); index >= 0; index--) {
            WordSection section = _document.Sections[index];
            if (!section.HasRunningStoryReference("headerReference", kind)) continue;
            return section.PageSetup.GetHeader(kind);
        }
        return null;
    }

    private WordFooter? ResolveFooter(WordHeaderFooterType kind) {
        var snapshot = WordHeaderFooterResolutionSnapshot.Current(_document);
        if (snapshot != null && snapshot.TryGet(this, out var stories)) return stories.Footer(kind);
        for (int index = GetSectionOrdinal(); index >= 0; index--) {
            WordSection section = _document.Sections[index];
            if (!section.HasRunningStoryReference("footerReference", kind)) continue;
            return section.PageSetup.GetFooter(kind);
        }
        return null;
    }
}
