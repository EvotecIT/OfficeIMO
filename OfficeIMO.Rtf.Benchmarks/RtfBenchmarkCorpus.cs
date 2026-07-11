using System.Collections.Concurrent;
using System.Text;
using OfficeIMO.Rtf;

namespace OfficeIMO.Rtf.Benchmarks;

internal sealed class RtfBenchmarkFixture {
    public RtfBenchmarkFixture(string scale, string rtf, int paragraphCount) {
        Scale = scale;
        Rtf = rtf;
        ParagraphCount = paragraphCount;
    }

    public string Scale { get; }
    public string Rtf { get; }
    public int ParagraphCount { get; }
    public int InputBytes => Encoding.UTF8.GetByteCount(Rtf);
}

internal static class RtfBenchmarkCorpus {
    private static readonly ConcurrentDictionary<string, RtfBenchmarkFixture> Fixtures =
        new ConcurrentDictionary<string, RtfBenchmarkFixture>(StringComparer.OrdinalIgnoreCase);

    public static IReadOnlyList<string> Scales { get; } = new[] { "Small", "Medium", "Large" };

    public static RtfBenchmarkFixture Get(string scale) {
        if (!Scales.Contains(scale, StringComparer.OrdinalIgnoreCase)) {
            throw new ArgumentException($"Unknown RTF benchmark scale '{scale}'.", nameof(scale));
        }

        return Fixtures.GetOrAdd(scale, Create);
    }

    private static RtfBenchmarkFixture Create(string scale) {
        int paragraphCount = string.Equals(scale, "Small", StringComparison.OrdinalIgnoreCase) ? 12 :
            string.Equals(scale, "Medium", StringComparison.OrdinalIgnoreCase) ? 250 : 2_000;
        RtfDocument document = RtfDocument.Create();
        int navy = document.AddColor(31, 78, 121);
        int gray = document.AddColor(89, 89, 89);
        document.AddHeader().AddParagraph("OfficeIMO RTF benchmark corpus");
        document.AddFooter().AddParagraph($"Scale: {scale}");

        for (int index = 0; index < paragraphCount; index++) {
            RtfParagraph paragraph = document.AddParagraph();
            paragraph.AddText($"Record {index + 1}: ").SetBold().ForegroundColorIndex = navy;
            paragraph.AddText("A representative clinical and business paragraph with Unicode ");
            paragraph.AddText(index % 2 == 0 ? "zażółć gęślą jaźń" : "Καλημέρα Привет").SetItalic();
            paragraph.AddText(" and deterministic fields for parsing, conversion, and extraction.")
                .ForegroundColorIndex = gray;

            if (index % 10 == 0) paragraph.SetList(1, index % 3, RtfListKind.Decimal);
            if (index > 0 && index % 50 == 0) paragraph.AddFootnote("1", $"Benchmark note for record {index + 1}.");
            if (index % 25 == 0) AddTable(document, index);
            if (index % 200 == 0) AddImage(document);
        }

        return new RtfBenchmarkFixture(scale, document.ToRtf(new RtfWriteOptions { IncludeGenerator = false }), paragraphCount);
    }

    private static void AddTable(RtfDocument document, int index) {
        RtfTable table = document.AddTable(3, 3);
        for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
            for (int columnIndex = 0; columnIndex < table.Rows[rowIndex].Cells.Count; columnIndex++) {
                table.Rows[rowIndex].Cells[columnIndex].AddParagraph($"R{index + rowIndex + 1} C{columnIndex + 1}");
            }
        }
    }

    private static void AddImage(RtfDocument document) {
        byte[] png = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");
        RtfImage image = document.AddImage(RtfImageFormat.Png, png);
        image.SourceWidth = 1;
        image.SourceHeight = 1;
        image.DesiredWidthTwips = 240;
        image.DesiredHeightTwips = 240;
    }
}
