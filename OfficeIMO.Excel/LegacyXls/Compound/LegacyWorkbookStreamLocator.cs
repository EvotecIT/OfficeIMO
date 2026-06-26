namespace OfficeIMO.Excel.LegacyXls.Compound {
    internal static class LegacyWorkbookStreamLocator {
        internal static byte[]? FindWorkbookStream(IReadOnlyDictionary<string, byte[]> streams) {
            if (streams.TryGetValue("Workbook", out byte[]? workbook)) {
                return workbook;
            }

            if (streams.TryGetValue("Book", out byte[]? book)) {
                return book;
            }

            foreach (var item in streams) {
                if (string.Equals(item.Key, "Workbook", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(item.Key, "Book", StringComparison.OrdinalIgnoreCase)) {
                    return item.Value;
                }
            }

            return null;
        }
    }
}
