namespace OfficeIMO.Pdf;

internal static class PdfInfoDictionaryBuilder {
    internal static string Build(string? title, string? author, string? subject, string? keywords, PdfTrappingStatus? trappingStatus = null) {
        var sb = new StringBuilder("<< ");
        AppendInfoString(sb, "Title", title);
        AppendInfoString(sb, "Author", author);
        AppendInfoString(sb, "Subject", subject);
        AppendInfoString(sb, "Keywords", keywords);
        sb.Append("/Producer (OfficeIMO.Pdf) ");
        AppendTrappingStatus(sb, trappingStatus);
        sb.Append(">>\n");
        return sb.ToString();
    }

    internal static string Build(PdfMetadata metadata) {
        Guard.NotNull(metadata, nameof(metadata));
        return Build(metadata.Title, metadata.Author, metadata.Subject, metadata.Keywords, metadata.TrappingStatus);
    }

    internal static PdfDictionary BuildDictionary(PdfMetadata metadata) {
        Guard.NotNull(metadata, nameof(metadata));
        var dictionary = new PdfDictionary();
        AddInfoString(dictionary, "Title", metadata.Title);
        AddInfoString(dictionary, "Author", metadata.Author);
        AddInfoString(dictionary, "Subject", metadata.Subject);
        AddInfoString(dictionary, "Keywords", metadata.Keywords);
        dictionary.Items["Producer"] = new PdfStringObj("OfficeIMO.Pdf");
        if (metadata.TrappingStatus.HasValue) {
            Guard.TrappingStatus(metadata.TrappingStatus.Value, nameof(metadata.TrappingStatus));
            dictionary.Items["Trapped"] = new PdfName(metadata.TrappingStatus.Value switch {
                PdfTrappingStatus.True => "True",
                PdfTrappingStatus.False => "False",
                _ => "Unknown"
            });
        }
        return dictionary;
    }

    private static void AppendInfoString(StringBuilder sb, string key, string? value) {
        if (string.IsNullOrEmpty(value)) {
            return;
        }

        sb.Append('/')
            .Append(PdfSyntaxEscaper.Name(key))
            .Append(' ')
            .Append(PdfSyntaxEscaper.TextString(value!))
            .Append(' ');
    }

    private static void AddInfoString(PdfDictionary dictionary, string key, string? value) {
        if (!string.IsNullOrEmpty(value)) {
            dictionary.Items[key] = new PdfStringObj(value!);
        }
    }

    private static void AppendTrappingStatus(StringBuilder sb, PdfTrappingStatus? trappingStatus) {
        if (!trappingStatus.HasValue) {
            return;
        }

        Guard.TrappingStatus(trappingStatus.Value, nameof(trappingStatus));
        sb.Append("/Trapped /")
            .Append(trappingStatus.Value switch {
                PdfTrappingStatus.True => "True",
                PdfTrappingStatus.False => "False",
                _ => "Unknown"
            })
            .Append(' ');
    }
}
