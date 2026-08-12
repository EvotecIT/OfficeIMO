using System.Globalization;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

public enum PdfBenchmarkScale {
    Easy,
    Medium,
    High
}

public enum PdfBenchmarkProducer {
    OfficeIMO,
    QuestPDF,
    PeachPDF,
    MigraDoc,
    IText
}

internal sealed record PdfBenchmarkScenario(
    PdfBenchmarkScale Scale,
    string Name,
    int PageCount,
    int RowsPerPage,
    int ParagraphsPerPage,
    int DocumentNumber = 0) {
    internal static PdfBenchmarkScenario Get(PdfBenchmarkScale scale) => scale switch {
        PdfBenchmarkScale.Easy => new(scale, "Invoice", 1, 8, 1),
        PdfBenchmarkScale.Medium => new(scale, "Monthly operational report", 20, 12, 2),
        PdfBenchmarkScale.High => new(scale, "Annual audit archive", 100, 12, 3),
        _ => throw new ArgumentOutOfRangeException(nameof(scale))
    };

    internal string PageTitle(int pageNumber) => DocumentNumber == 0
        ? $"Benchmark Report {pageNumber.ToString("D3", CultureInfo.InvariantCulture)}"
        : $"Benchmark Report {DocumentNumber.ToString("D2", CultureInfo.InvariantCulture)}-{pageNumber.ToString("D3", CultureInfo.InvariantCulture)}";

    internal string Narrative(int pageNumber, int paragraphNumber) =>
        $"Operational summary {paragraphNumber + 1} for page {pageNumber:D3}. " +
        "This report records deterministic account, owner, amount, status, and review data.";

    internal IReadOnlyList<string[]> TableRows(int pageNumber) {
        var rows = new List<string[]>(RowsPerPage + 1) {
            new[] { "Account", "Owner", "Amount", "Status" }
        };
        for (int row = 1; row <= RowsPerPage; row++) {
            rows.Add(new[] {
                $"ACC-{pageNumber:D3}-{row:D2}",
                $"Owner {(pageNumber + row) % 17:D2}",
                ((pageNumber * 1000M) + (row * 37.25M)).ToString("0.00", CultureInfo.InvariantCulture),
                row % 4 == 0 ? "Review" : "Approved"
            });
        }

        return rows;
    }
}

internal static class PdfBenchmarkAssets {
    private static readonly Lazy<byte[]> Font = new(
        static () => File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fonts", "Carlito-Regular.ttf")),
        isThreadSafe: true);

    internal static byte[] CarlitoRegular => Font.Value;
}
