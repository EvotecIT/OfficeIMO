namespace OfficeIMO.Pdf;

internal static class PdfInfoDictionaryBuilder {
    internal static string Build(
        string? title,
        string? author,
        string? subject,
        string? keywords,
        PdfTrappingStatus? trappingStatus = null,
        PdfXIdentification? pdfXIdentification = null,
        PdfXProductionMetadata? pdfXProductionMetadata = null) =>
        BuildCore(
            title,
            author,
            subject,
            keywords,
            trappingStatus,
            pdfXProductionMetadata?.CreationDate,
            pdfXProductionMetadata?.ModificationDate,
            pdfXIdentification?.Version,
            pdfXIdentification?.Conformance);

    private static string BuildCore(
        string? title,
        string? author,
        string? subject,
        string? keywords,
        PdfTrappingStatus? trappingStatus,
        DateTimeOffset? creationDate,
        DateTimeOffset? modificationDate,
        string? pdfXVersion,
        string? pdfXConformance) {
        var sb = new StringBuilder("<< ");
        AppendInfoString(sb, "Title", title);
        AppendInfoString(sb, "Author", author);
        AppendInfoString(sb, "Subject", subject);
        AppendInfoString(sb, "Keywords", keywords);
        sb.Append("/Producer (OfficeIMO.Pdf) ");
        AppendProductionMetadata(sb, creationDate, modificationDate, pdfXVersion, pdfXConformance);
        AppendTrappingStatus(sb, trappingStatus);
        sb.Append(">>\n");
        return sb.ToString();
    }

    internal static string Build(PdfMetadata metadata) {
        Guard.NotNull(metadata, nameof(metadata));
        var sb = new StringBuilder("<< ");
        AppendInfoString(sb, "Title", metadata.Title);
        AppendInfoString(sb, "Author", metadata.Author);
        AppendInfoString(sb, "Subject", metadata.Subject);
        AppendInfoString(sb, "Keywords", metadata.Keywords);
        sb.Append("/Producer (OfficeIMO.Pdf) ");
        AppendInfoString(sb, "CreationDate", metadata.CreationDateRaw ??
            (metadata.CreationDate.HasValue ? PdfDateCodec.Format(metadata.CreationDate.Value) : null));
        AppendInfoString(sb, "ModDate", metadata.ModificationDateRaw ??
            (metadata.ModificationDate.HasValue ? PdfDateCodec.Format(metadata.ModificationDate.Value) : null));
        AppendInfoString(sb, "GTS_PDFXVersion", metadata.PdfXVersion);
        AppendInfoString(sb, "GTS_PDFXConformance", metadata.PdfXConformance);
        AppendTrappingStatus(sb, metadata.TrappingStatus);
        sb.Append(">>\n");
        return sb.ToString();
    }

    internal static PdfDictionary BuildDictionary(PdfMetadata metadata) {
        Guard.NotNull(metadata, nameof(metadata));
        var dictionary = new PdfDictionary();
        AddInfoString(dictionary, "Title", metadata.Title);
        AddInfoString(dictionary, "Author", metadata.Author);
        AddInfoString(dictionary, "Subject", metadata.Subject);
        AddInfoString(dictionary, "Keywords", metadata.Keywords);
        dictionary.Items["Producer"] = new PdfStringObj("OfficeIMO.Pdf");
        AddInfoString(dictionary, "CreationDate", metadata.CreationDateRaw ??
            (metadata.CreationDate.HasValue ? PdfDateCodec.Format(metadata.CreationDate.Value) : null));
        AddInfoString(dictionary, "ModDate", metadata.ModificationDateRaw ??
            (metadata.ModificationDate.HasValue ? PdfDateCodec.Format(metadata.ModificationDate.Value) : null));
        AddInfoString(dictionary, "GTS_PDFXVersion", metadata.PdfXVersion);
        AddInfoString(dictionary, "GTS_PDFXConformance", metadata.PdfXConformance);
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

    private static void AppendProductionMetadata(
        StringBuilder sb,
        DateTimeOffset? creationDate,
        DateTimeOffset? modificationDate,
        string? pdfXVersion,
        string? pdfXConformance) {
        AppendInfoString(sb, "CreationDate", creationDate.HasValue ? PdfDateCodec.Format(creationDate.Value) : null);
        AppendInfoString(sb, "ModDate", modificationDate.HasValue ? PdfDateCodec.Format(modificationDate.Value) : null);
        AppendInfoString(sb, "GTS_PDFXVersion", pdfXVersion);
        AppendInfoString(sb, "GTS_PDFXConformance", pdfXConformance);
    }

}
