namespace OfficeIMO.Pdf;

public sealed partial class PdfReadDocument {
    private string ToRaw() {
        // Reconstruct raw text for simple metadata extraction without reserialization; ok for small files.
        var sb = new StringBuilder();
        foreach (var kv in _objects.OrderBy(k => k.Key)) {
            sb.Append(kv.Key).Append(" 0 obj\n");
            if (kv.Value.Value is PdfStream s) {
                sb.Append("<< ");
                foreach (var d in s.Dictionary.Items) sb.Append('/').Append(d.Key).Append(' ').Append(' ').Append(' ');
                sb.Append(">>\nstream\n");
                sb.Append(PdfEncoding.Latin1GetString(s.Data)).Append("\nendstream\nendobj\n");
            } else {
                sb.Append("...\nendobj\n");
            }
        }
        sb.Append(_trailerRaw);
        return sb.ToString();
    }

    private PdfMetadata ExtractMetadata() {
        // Trailer has /Info N G R when present.
        if (!PdfSyntax.TryGetTrailerReference(_trailerRaw, "Info", _options.Limits, out PdfReference infoReference)) {
            return new PdfMetadata();
        }
        if (!PdfObjectLookup.TryGet(_objects, infoReference, out var infoObj) ||
            infoObj.Value is not PdfDictionary dict) return new PdfMetadata();
        string? GetText(string key) =>
            dict.Items.TryGetValue(key, out PdfObject? value) && ResolveObject(value) is PdfStringObj text
                ? text.Value
                : null;
        string? GetName(string key) =>
            dict.Items.TryGetValue(key, out PdfObject? value) && ResolveObject(value) is PdfName name
                ? name.Name
                : null;
        return new PdfMetadata {
            Title = GetText("Title"),
            Author = GetText("Author"),
            Subject = GetText("Subject"),
            Keywords = GetText("Keywords"),
            TrappingStatus = ParseTrappingStatus(GetName("Trapped")),
            CreationDate = PdfDateCodec.TryParse(GetText("CreationDate")),
            ModificationDate = PdfDateCodec.TryParse(GetText("ModDate")),
            PdfXVersion = GetText("GTS_PDFXVersion"),
            PdfXConformance = GetText("GTS_PDFXConformance")
        };
    }

    private static PdfTrappingStatus? ParseTrappingStatus(string? value) =>
        value switch {
            "True" => PdfTrappingStatus.True,
            "False" => PdfTrappingStatus.False,
            "Unknown" => PdfTrappingStatus.Unknown,
            _ => null
        };
}
