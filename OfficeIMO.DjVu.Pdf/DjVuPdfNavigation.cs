using OfficeIMO.Pdf;

namespace OfficeIMO.DjVu.Pdf;

internal static class DjVuPdfNavigation {
    internal sealed class Entry {
        internal string Title = string.Empty;
        internal string? Uri;
        internal int SourcePage, Level, Order;
    }
    internal static IReadOnlyList<Entry> Project(DjVuDocument document, DjVuPage[] pages,
        List<OfficeConversionFidelityDiagnostic> diagnostics, CancellationToken token) {
        var selected = new HashSet<int>(pages.Select(p => p.Number));
        var output = new List<Entry>();
        if (document.BookmarkDiagnostic != null) DjVuPdfConversionEngine.Add(diagnostics, "djvu.pdf.outline-corrupt",
            document.BookmarkDiagnostic, OfficeConversionLossKind.Omission, null);
        foreach (var bookmark in document.Bookmarks) Visit(bookmark, 1);
        return output.AsReadOnly();
        int? FirstDestination(DjVuBookmark bookmark) {
            token.ThrowIfCancellationRequested();
            if (bookmark.PageNumber.HasValue && selected.Contains(bookmark.PageNumber.Value)) return bookmark.PageNumber;
            foreach (var child in bookmark.Children) {
                int? target = FirstDestination(child);
                if (target.HasValue) return target;
            }
            return null;
        }
        void Visit(DjVuBookmark bookmark, int level) {
            token.ThrowIfCancellationRequested();
            string? uri = null;
            int? page = bookmark.PageNumber;
            bool group = bookmark.Target.Length == 0;
            bool usable = !string.IsNullOrWhiteSpace(bookmark.Title) && level <= 64;
            if (group) page = FirstDestination(bookmark) ?? pages[0].Number;
            else if (page.HasValue) usable &= selected.Contains(page.Value);
            else {
                usable &= Guard.IsUriAction(bookmark.Target, token);
                // An unresolved local component target stays a reported omission;
                // it must not turn into a new external file/URL action in the PDF.
                usable &= Uri.TryCreate(bookmark.Target, UriKind.Absolute, out _);
                uri = bookmark.Target;
                page = pages[0].Number;
            }
            if (usable) {
                output.Add(new Entry { Title = bookmark.Title, Uri = uri, SourcePage = page!.Value, Level = level, Order = output.Count });
                if (group) DjVuPdfConversionEngine.Add(diagnostics, "djvu.pdf.outline-group-anchor",
                    "A grouping bookmark is anchored to its first exported descendant, or the first exported page when it has no page target.", OfficeConversionLossKind.Approximation, null);
            } else DjVuPdfConversionEngine.Add(diagnostics, "djvu.pdf.outline-entry-omitted",
                "A bookmark has an unavailable destination, unusable title or unsupported outline depth.", OfficeConversionLossKind.Omission, bookmark.PageNumber);
            foreach (var child in bookmark.Children) Visit(child, usable ? level + 1 : level);
        }
    }
}
