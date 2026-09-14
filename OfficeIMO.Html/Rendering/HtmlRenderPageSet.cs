namespace OfficeIMO.Html;

/// <summary>Immutable selection and composition policy for rendered HTML pages.</summary>
public sealed class HtmlRenderPageSet {
    private HtmlRenderPageSet(HtmlRenderPageSetMode mode, int firstPageIndex, int? pageCount) {
        Mode = mode;
        FirstPageIndex = firstPageIndex;
        PageCount = pageCount;
    }

    /// <summary>How selected pages are exposed to the output adapter.</summary>
    public HtmlRenderPageSetMode Mode { get; }

    /// <summary>Zero-based first source page included by the selection.</summary>
    public int FirstPageIndex { get; }

    /// <summary>Maximum pages included by a range, or <see langword="null"/> for every remaining page.</summary>
    public int? PageCount { get; }

    /// <summary>Returns every rendered page as a separate ordered surface.</summary>
    public static HtmlRenderPageSet All() => new(HtmlRenderPageSetMode.Separate, 0, null);

    /// <summary>Returns one selected zero-based page.</summary>
    public static HtmlRenderPageSet Page(int pageIndex) {
        if (pageIndex < 0) throw new ArgumentOutOfRangeException(nameof(pageIndex));
        return new HtmlRenderPageSet(HtmlRenderPageSetMode.Selected, pageIndex, 1);
    }

    /// <summary>Returns a bounded range of pages as separate ordered surfaces.</summary>
    public static HtmlRenderPageSet Pages(int firstPageIndex, int pageCount) {
        if (firstPageIndex < 0) throw new ArgumentOutOfRangeException(nameof(firstPageIndex));
        if (pageCount <= 0) throw new ArgumentOutOfRangeException(nameof(pageCount));
        return new HtmlRenderPageSet(HtmlRenderPageSetMode.Range, firstPageIndex, pageCount);
    }

    /// <summary>Places every rendered page vertically on one continuous surface.</summary>
    public static HtmlRenderPageSet Stitched() => new(HtmlRenderPageSetMode.Stitched, 0, null);

    /// <summary>Selects archive-plus-manifest packaging for separate encoded pages.</summary>
    public static HtmlRenderPageSet Archive() => new(HtmlRenderPageSetMode.ArchiveWithManifest, 0, null);

    internal IReadOnlyList<HtmlRenderPage> Select(IReadOnlyList<HtmlRenderPage> pages) {
        if (pages == null) throw new ArgumentNullException(nameof(pages));
        if (pages.Count == 0) throw new ArgumentException("At least one rendered page is required.", nameof(pages));
        if (Mode == HtmlRenderPageSetMode.Separate || Mode == HtmlRenderPageSetMode.Stitched || Mode == HtmlRenderPageSetMode.ArchiveWithManifest) {
            return pages;
        }
        if (FirstPageIndex >= pages.Count) {
            throw new ArgumentOutOfRangeException(nameof(FirstPageIndex), "The selected HTML render page does not exist.");
        }
        int count = Math.Min(PageCount ?? pages.Count, pages.Count - FirstPageIndex);
        return pages.Skip(FirstPageIndex).Take(count).ToList().AsReadOnly();
    }

    internal IReadOnlyList<int> SelectIndices(int pageCount) {
        if (pageCount <= 0) throw new ArgumentOutOfRangeException(nameof(pageCount));
        if (Mode == HtmlRenderPageSetMode.Separate || Mode == HtmlRenderPageSetMode.Stitched || Mode == HtmlRenderPageSetMode.ArchiveWithManifest) {
            return Enumerable.Range(0, pageCount).ToList().AsReadOnly();
        }
        if (FirstPageIndex >= pageCount) {
            throw new ArgumentOutOfRangeException(nameof(FirstPageIndex), "The selected HTML render page does not exist.");
        }
        int count = Math.Min(PageCount ?? pageCount, pageCount - FirstPageIndex);
        return Enumerable.Range(FirstPageIndex, count).ToList().AsReadOnly();
    }
}
