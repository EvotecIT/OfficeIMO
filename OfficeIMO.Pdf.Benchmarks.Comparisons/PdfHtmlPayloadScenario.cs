using System.Text;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

public enum PdfHtmlPayloadKind {
    PlainText,
    Table,
    Multilingual
}

internal sealed record PdfHtmlPayloadScenario(
    PdfHtmlPayloadKind Kind,
    string Html,
    IReadOnlyList<string> RequiredText) {
    internal const int TargetUtf8Bytes = 21 * 1024;

    internal static PdfHtmlPayloadScenario Create(PdfHtmlPayloadKind kind) {
        string prefix = kind == PdfHtmlPayloadKind.Table
            ? "<!doctype html><html><head><meta charset='utf-8'><style>@page{size:A4;margin:12mm}body{font:10pt sans-serif}table{width:100%;border-collapse:collapse}td,th{border:1px solid #789;padding:3px}</style></head><body><h1>HTML-PDF-21K</h1><table><thead><tr><th>Id</th><th>Description</th><th>Amount</th></tr></thead><tbody>"
            : "<!doctype html><html><head><meta charset='utf-8'><style>@page{size:A4;margin:12mm}body{font:10pt sans-serif}p{margin:0 0 6pt}</style></head><body><h1>HTML-PDF-21K</h1>";
        string suffix = kind == PdfHtmlPayloadKind.Table ? "</tbody></table></body></html>" : "</body></html>";
        string unit = kind switch {
            PdfHtmlPayloadKind.PlainText => "<p>ITEM-0001 · Deterministic conversion evidence for account records, pagination, searchable text, and output size.</p>",
            PdfHtmlPayloadKind.Table => "<tr><td>ROW-0001</td><td>Deterministic account conversion evidence</td><td>1234.50</td></tr>",
            PdfHtmlPayloadKind.Multilingual => "<p>ITEM-0001 · Zażółć gęślą jaźń · Ελληνικά · Русский · العربية · עברית · 日本語 · HTML PDF evidence.</p>",
            _ => throw new ArgumentOutOfRangeException(nameof(kind))
        };

        var html = new StringBuilder(TargetUtf8Bytes + 64).Append(prefix);
        int index = 0;
        int currentBytes = Encoding.UTF8.GetByteCount(prefix);
        int suffixBytes = Encoding.UTF8.GetByteCount(suffix);
        while (true) {
            string next = unit.Replace("0001", (index + 1).ToString("D4"), StringComparison.Ordinal);
            int nextBytes = Encoding.UTF8.GetByteCount(next);
            int projectedBytes = currentBytes
                + nextBytes
                + suffixBytes
                + 7;
            if (projectedBytes > TargetUtf8Bytes) break;
            html.Append(next);
            currentBytes += nextBytes;
            index++;
        }

        html.Append(suffix);
        int remaining = TargetUtf8Bytes - Encoding.UTF8.GetByteCount(html.ToString());
        if (remaining < 7) {
            throw new InvalidOperationException("The HTML payload padding budget is too small.");
        }

        html.Insert(html.Length - suffix.Length, "<!--" + new string('x', remaining - 7) + "-->");
        string value = html.ToString();
        int actualBytes = Encoding.UTF8.GetByteCount(value);
        if (actualBytes != TargetUtf8Bytes) {
            throw new InvalidOperationException($"Expected {TargetUtf8Bytes} UTF-8 bytes but generated {actualBytes}.");
        }

        return new PdfHtmlPayloadScenario(
            kind,
            value,
            kind switch {
                PdfHtmlPayloadKind.PlainText => new[] {
                    "HTML-PDF-21K",
                    "ITEM-0001",
                    $"ITEM-{index:D4}",
                    "Deterministic conversion evidence"
                },
                PdfHtmlPayloadKind.Table => new[] { "HTML-PDF-21K", "ROW-0001", $"ROW-{index:D4}" },
                PdfHtmlPayloadKind.Multilingual => new[] {
                    "HTML-PDF-21K",
                    "ITEM-0001",
                    $"ITEM-{index:D4}",
                    "Zażółć gęślą jaźń",
                    "Ελληνικά",
                    "Русский",
                    "HTML PDF evidence"
                },
                _ => throw new ArgumentOutOfRangeException(nameof(kind))
            });
    }
}
