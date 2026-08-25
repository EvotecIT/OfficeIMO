using System.Collections.Concurrent;

namespace OfficeIMO.Rtf.Benchmarks.Comparisons;

internal sealed record RtfHtmlComparisonFixture(string Scale, string Rtf, int ParagraphCount);

internal static class RtfHtmlComparisonCorpus {
    private static readonly ConcurrentDictionary<string, RtfHtmlComparisonFixture> Fixtures =
        new(StringComparer.OrdinalIgnoreCase);

    internal static IReadOnlyList<string> Scales { get; } = ["Small", "Medium", "Large", "Producer"];

    internal static RtfHtmlComparisonFixture Get(string scale) {
        if (!Scales.Contains(scale, StringComparer.OrdinalIgnoreCase)) {
            throw new ArgumentException($"Unknown RTF comparison scale '{scale}'.", nameof(scale));
        }

        return Fixtures.GetOrAdd(scale, Create);
    }

    private static RtfHtmlComparisonFixture Create(string scale) {
        if (string.Equals(scale, "Producer", StringComparison.OrdinalIgnoreCase)) {
            string path = Path.Combine(AppContext.BaseDirectory, "producer-gembox-document.rtf");
            return new RtfHtmlComparisonFixture(scale, File.ReadAllText(path), 3);
        }

        int paragraphCount = string.Equals(scale, "Small", StringComparison.OrdinalIgnoreCase)
            ? 12
            : string.Equals(scale, "Medium", StringComparison.OrdinalIgnoreCase) ? 250 : 2_000;
        RtfDocument document = RtfDocument.Create();
        int accent = document.AddColor(31, 78, 121);
        for (int index = 0; index < paragraphCount; index++) {
            RtfParagraph paragraph = document.AddParagraph();
            paragraph.AddText($"Record {index + 1}: ").SetBold().ForegroundColorIndex = accent;
            paragraph.AddText("A deterministic RTF-to-HTML comparison with Unicode ");
            paragraph.AddText(index % 2 == 0 ? "zażółć gęślą jaźń" : "Καλημέρα Привет").SetItalic();
            paragraph.AddText(" and stable paragraph content.");
        }

        RtfTable table = document.AddTable(3, 3);
        for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
            for (int columnIndex = 0; columnIndex < table.Rows[rowIndex].Cells.Count; columnIndex++) {
                table.Rows[rowIndex].Cells[columnIndex].AddParagraph($"R{rowIndex + 1} C{columnIndex + 1}");
            }
        }

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        return new RtfHtmlComparisonFixture(scale, rtf, paragraphCount);
    }
}
