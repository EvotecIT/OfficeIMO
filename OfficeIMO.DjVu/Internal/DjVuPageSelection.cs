namespace OfficeIMO.DjVu;

internal static class DjVuPageSelection {
    internal static DjVuPage[] Select(DjVuDocument document, IReadOnlyList<int>? numbers, int maximum, CancellationToken token) {
        int count = numbers?.Count ?? document.Pages.Count;
        if (count > maximum) throw new DjVuResourceLimitException("MaxSelectedPages");
        var pages = new DjVuPage[count];
        var seen = new HashSet<int>();
        for (int i = 0; i < count; i++) {
            token.ThrowIfCancellationRequested();
            int number = numbers?[i] ?? i + 1;
            if (number <= 0 || number > document.Pages.Count || !seen.Add(number))
                throw new ArgumentException("Page numbers must identify existing pages exactly once.", nameof(numbers));
            pages[i] = document.Pages[number - 1];
        }
        return pages;
    }
}
