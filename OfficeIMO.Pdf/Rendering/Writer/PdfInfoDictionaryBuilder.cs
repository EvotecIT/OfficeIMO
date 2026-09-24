using System.Threading;

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

    internal static string Build(PdfMetadata metadata) => BuildMetadataBuilder(metadata, CancellationToken.None).ToString();

    internal static byte[] BuildBytesCancellable(PdfMetadata metadata, CancellationToken cancellationToken) =>
        PdfEncoding.Latin1GetBytesCancellable(BuildMetadataBuilder(metadata, cancellationToken), cancellationToken);

    private static StringBuilder BuildMetadataBuilder(PdfMetadata metadata, CancellationToken cancellationToken) {
        Guard.NotNull(metadata, nameof(metadata));
        cancellationToken.ThrowIfCancellationRequested();
        var sb = new StringBuilder("<< ");
        AppendInfoString(sb, "Title", metadata.Title, cancellationToken);
        AppendInfoString(sb, "Author", metadata.Author, cancellationToken);
        AppendInfoString(sb, "Subject", metadata.Subject, cancellationToken);
        AppendInfoString(sb, "Keywords", metadata.Keywords, cancellationToken);
        sb.Append("/Producer (OfficeIMO.Pdf) ");
        AppendInfoString(sb, "CreationDate", metadata.CreationDateRaw ??
            (metadata.CreationDate.HasValue ? PdfDateCodec.Format(metadata.CreationDate.Value) : null), cancellationToken);
        AppendInfoString(sb, "ModDate", metadata.ModificationDateRaw ??
            (metadata.ModificationDate.HasValue ? PdfDateCodec.Format(metadata.ModificationDate.Value) : null), cancellationToken);
        AppendInfoString(sb, "GTS_PDFXVersion", metadata.PdfXVersion, cancellationToken);
        AppendInfoString(sb, "GTS_PDFXConformance", metadata.PdfXConformance, cancellationToken);
        AppendTrappingStatus(sb, metadata.TrappingStatus);
        sb.Append(">>\n");
        cancellationToken.ThrowIfCancellationRequested();
        return sb;
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

    private static void AppendInfoString(StringBuilder sb, string key, string? value, CancellationToken cancellationToken = default) {
        if (string.IsNullOrEmpty(value)) {
            return;
        }

        sb.Append('/').Append(PdfSyntaxEscaper.Name(key)).Append(' ');
        PdfSyntaxEscaper.AppendTextStringCancellable(sb, value!, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        sb.Append(' ');
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
