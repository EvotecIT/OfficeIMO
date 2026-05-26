namespace OfficeIMO.Pdf;

internal static class PdfCatalogDictionaryBuilder {
    internal static string BuildGeneratedCatalogDictionary(int pagesId, int outlinesId, int namedDestinationsId = 0) {
        var sb = new StringBuilder();
        AppendCatalogStart(sb, pagesId);

        if (outlinesId < 0) {
            throw new ArgumentOutOfRangeException(nameof(outlinesId), "PDF outline object number cannot be negative.");
        }

        if (namedDestinationsId < 0) {
            throw new ArgumentOutOfRangeException(nameof(namedDestinationsId), "PDF named destinations object number cannot be negative.");
        }

        if (outlinesId > 0) {
            AppendReferenceEntry(sb, "Outlines", outlinesId);
            AppendNameEntry(sb, "PageMode", "UseOutlines");
        }

        if (namedDestinationsId > 0) {
            AppendNamedDestinationsEntry(sb, namedDestinationsId);
        }

        sb.Append(" >>\n");
        return sb.ToString();
    }

    internal static void AppendCatalogStart(StringBuilder sb, int pagesId) {
        Guard.NotNull(sb, nameof(sb));
        sb.Append("<< /Type /Catalog /Pages ")
            .Append(PdfSyntaxEscaper.IndirectReference(pagesId));
    }

    internal static void AppendNameEntry(StringBuilder sb, string key, string value) {
        Guard.NotNull(sb, nameof(sb));
        Guard.NotNullOrWhiteSpace(key, nameof(key));
        Guard.NotNullOrWhiteSpace(value, nameof(value));
        sb.Append(" /")
            .Append(PdfSyntaxEscaper.Name(key))
            .Append(" /")
            .Append(PdfSyntaxEscaper.Name(value));
    }

    internal static void AppendReferenceEntry(StringBuilder sb, string key, int objectNumber, int generation = 0) {
        Guard.NotNull(sb, nameof(sb));
        Guard.NotNullOrWhiteSpace(key, nameof(key));
        sb.Append(" /")
            .Append(PdfSyntaxEscaper.Name(key))
            .Append(' ')
            .Append(PdfSyntaxEscaper.IndirectReference(objectNumber, generation));
    }

    private static void AppendNamedDestinationsEntry(StringBuilder sb, int objectNumber) {
        sb.Append(" /Names << /Dests ")
            .Append(PdfSyntaxEscaper.IndirectReference(objectNumber))
            .Append(" >>");
    }
}
