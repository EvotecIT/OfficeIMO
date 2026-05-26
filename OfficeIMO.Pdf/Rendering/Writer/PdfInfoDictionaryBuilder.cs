namespace OfficeIMO.Pdf;

internal static class PdfInfoDictionaryBuilder {
    internal static string Build(string? title, string? author, string? subject, string? keywords) {
        var sb = new StringBuilder("<< ");
        AppendInfoString(sb, "Title", title);
        AppendInfoString(sb, "Author", author);
        AppendInfoString(sb, "Subject", subject);
        AppendInfoString(sb, "Keywords", keywords);
        sb.Append("/Producer (OfficeIMO.Pdf) >>\n");
        return sb.ToString();
    }

    internal static string Build(PdfMetadata metadata) {
        Guard.NotNull(metadata, nameof(metadata));
        return Build(metadata.Title, metadata.Author, metadata.Subject, metadata.Keywords);
    }

    private static void AppendInfoString(StringBuilder sb, string key, string? value) {
        if (string.IsNullOrEmpty(value)) {
            return;
        }

        sb.Append('/')
            .Append(PdfSyntaxEscaper.Name(key))
            .Append(' ')
            .Append(PdfSyntaxEscaper.LiteralString(value!))
            .Append(' ');
    }
}
