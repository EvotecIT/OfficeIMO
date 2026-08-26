namespace OfficeIMO.Pdf;

public sealed partial class PdfReadDocument {
    private static readonly System.Text.RegularExpressions.Regex InfoReferenceRegex =
        new System.Text.RegularExpressions.Regex(
            @"/Info\s+(\d+)\s+(\d+)\s+R",
            System.Text.RegularExpressions.RegexOptions.Compiled
                | System.Text.RegularExpressions.RegexOptions.CultureInvariant,
            TimeSpan.FromSeconds(2));

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
        System.Text.RegularExpressions.Match m = InfoReferenceRegex.Match(_trailerRaw);
        if (!m.Success) return new PdfMetadata();
        if (!int.TryParse(m.Groups[1].Value, out int infoId)) return new PdfMetadata();
        if (!int.TryParse(m.Groups[2].Value, out int infoGeneration)) return new PdfMetadata();
        if (!PdfObjectLookup.TryGet(_objects, new PdfReference(infoId, infoGeneration), out var infoObj) ||
            infoObj.Value is not PdfDictionary dict) return new PdfMetadata();
        string? GetStr(string key) => dict.Get<PdfStringObj>(key)?.Value;
        return new PdfMetadata {
            Title = GetStr("Title"),
            Author = GetStr("Author"),
            Subject = GetStr("Subject"),
            Keywords = GetStr("Keywords"),
            TrappingStatus = ParseTrappingStatus(dict.Get<PdfName>("Trapped")?.Name),
            CreationDate = PdfDateCodec.TryParse(GetStr("CreationDate")),
            ModificationDate = PdfDateCodec.TryParse(GetStr("ModDate")),
            PdfXVersion = GetStr("GTS_PDFXVersion"),
            PdfXConformance = GetStr("GTS_PDFXConformance")
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
