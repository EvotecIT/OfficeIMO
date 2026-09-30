using System.Security.Cryptography;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using OfficeIMO.Ocr;
using OfficeIMO.Ocr.Tesseract;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderOcrIndependentLayoutTests {
    [Theory]
    [InlineData("english-columns")]
    [InlineData("english-ledger")]
    public async Task IndependentProducerLayoutSurvivesNativeOcrAndSearchableReadback(string id) {
        string root = Path.Combine(AppContext.BaseDirectory, "EnglishLayout");
        using JsonDocument manifest = JsonDocument.Parse(File.ReadAllText(Path.Combine(root, "manifest.json")));
        JsonElement fixture = manifest.RootElement.GetProperty("cases").EnumerateArray()
            .Single(item => item.GetProperty("id").GetString() == id);
        string[] expected = fixture.GetProperty("readingOrder").EnumerateArray().Select(item => item.GetString()!).ToArray();
        string[][] table = fixture.GetProperty("table").EnumerateArray()
            .Select(row => row.EnumerateArray().Select(item => item.GetString()!).ToArray()).ToArray();
        byte[] native = ReadVerified("native.pdf"), scan = ReadVerified("scan.pdf");
        string tsv = System.Text.Encoding.UTF8.GetString(ReadVerified("provider.tsv"));
        string[] pageRow = tsv.Split('\n').Select(line => line.TrimEnd('\r').Split('\t'))
            .First(columns => columns.Length >= 10 && columns[0] == "1");
        double width = double.Parse(pageRow[8], System.Globalization.CultureInfo.InvariantCulture);
        double height = double.Parse(pageRow[9], System.Globalization.CultureInfo.InvariantCulture);
        var engine = new DelegateOcrEngine("recorded-tesseract-layout", (_, _) => {
            OcrResult result = TesseractTsvParser.Parse(tsv, "eng");
            // Replay independently recorded provider geometry; normalized coordinates preserve
            // its measured boxes across the consumer's selected raster resolution.
            foreach (OcrTextSpan span in result.Spans) {
                span.CoordinateUnit = OcrCoordinateUnit.Normalized;
                span.Region!.X /= width;
                span.Region.Y /= height;
                span.Region.Width /= width;
                span.Region.Height /= height;
            }
            return Task.FromResult(result);
        }, new OcrEngineCapabilities { SupportedMediaTypes = new[] { "image/png" }, SupportsWordSpans = true });
        AssertLayout(PdfDocument.Load(native).Read(new PdfReadOptions { Profile = PdfReadProfile.Structured }));
        PdfDocument source = PdfDocument.Load(scan);
        PdfSearchableOcrResult searchable = await source.MakeSearchableAsync(engine,
            new PdfOcrMergeOptions { Dpi = 150, MinimumConfidence = 0, ReconstructLayout = true });
        AssertLayout(searchable.Ocr.Document);
        AssertLayout(PdfDocument.Load(searchable.Document.ToBytes()).Read(new PdfReadOptions { Profile = PdfReadProfile.Structured }));
        Assert.Equal(scan, source.ToBytes());

        byte[] ReadVerified(string suffix) {
            byte[] bytes = File.ReadAllBytes(Path.Combine(root, id + "-" + suffix));
            using SHA256 hash = SHA256.Create();
            string actual = string.Concat(hash.ComputeHash(bytes).Select(static value => value.ToString("x2")));
            Assert.Equal(fixture.GetProperty("files").GetProperty(suffix).GetString(), actual);
            return bytes;
        }

        void AssertLayout(PdfDocumentReadResult document) {
            Assert.Equal(Normalize(string.Join(" ", expected)), Normalize(document.Text));
            PdfLogicalTable[] detected = document.Tables.ToArray();
            if (table.Length == 0) Assert.Empty(detected);
            else {
                PdfLogicalTable actual = Assert.Single(detected);
                Assert.Equal(table.Length, actual.Rows.Count);
                for (int row = 0; row < table.Length; row++) Assert.Equal(table[row], actual.Rows[row]);
            }
        }
    }

    private static string Normalize(string text) => Regex.Replace(text.Normalize(), @"\s+", " ").Trim();
}
