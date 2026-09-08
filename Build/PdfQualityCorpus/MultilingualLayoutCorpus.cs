using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using OfficeIMO.Drawing;
using OfficeIMO.Ocr.Tesseract;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;

namespace OfficeIMO.PdfQualityCorpus;

internal static partial class MultilingualLayoutCorpus {
    internal static async Task<int> RunAsync(string[] args) {
        if (args.Length is < 3 or > 4)
            throw new ArgumentException("Usage: layout <MultilingualLayout fixtures> <output> [native]");
        string root = Path.GetFullPath(args[1]), output = Path.GetFullPath(args[2]);
        Directory.CreateDirectory(output);
        using var deadline = new CancellationTokenSource(TimeSpan.FromMinutes(15));
        using JsonDocument manifest = JsonDocument.Parse(await File.ReadAllTextAsync(Path.Combine(root, "manifest.json"), deadline.Token));
        var results = new List<object>();
        var recognition = new List<object>();
        foreach (JsonElement fixture in manifest.RootElement.GetProperty("cases").EnumerateArray()) {
            string id = fixture.GetProperty("id").GetString()!;
            string[] expected = fixture.GetProperty("readingOrder").EnumerateArray().Select(item => item.GetString()!).ToArray();
            string[][] expectedTable = fixture.GetProperty("table").EnumerateArray()
                .Select(row => row.EnumerateArray().Select(item => item.GetString()!).ToArray()).ToArray();
            string caption = fixture.GetProperty("caption").GetString()!;
            byte[] native = await ReadVerifiedAsync(root, fixture, "native.pdf", deadline.Token);
            PdfDocumentReadResult nativeRead = PdfDocument.Load(native).Read(new PdfReadOptions { Profile = PdfReadProfile.Structured });
            await RecordAsync(id, "native", nativeRead, expected, expectedTable, caption, output, results, deadline.Token);
            if (args.Length == 4 && args[3] == "native") continue;

            byte[] scan = await ReadVerifiedAsync(root, fixture, "scan.pdf", deadline.Token);
            var engine = new TesseractOcrEngine(new TesseractOcrEngineOptions {
                Dpi = 300, Language = fixture.GetProperty("languages").GetString()!, PageSegmentationMode = 3,
                Timeout = TimeSpan.FromSeconds(45), TemporaryDirectory = Path.Combine(output, "temporary")
            });
            int turn = fixture.GetProperty("clockwiseDegrees").GetInt32() / 90;
            PdfSearchableOcrResult searchable = await PdfDocument.Load(scan).MakeSearchableAsync(engine,
                new PdfOcrMergeOptions {
#if !PDF_LAYOUT_BASELINE
                    ReconstructLayout = true,
#endif
                    Dpi = 300, MinimumConfidence = 0, ProviderTimeout = TimeSpan.FromSeconds(50),
                    // The independently labelled rotation is explicit: this lane measures layout,
                    // while the scan-quality corpus measures the provider's orientation detector.
                    ScanProcessing = new OfficeScanProcessingOptions { ClockwiseQuarterTurns = (4 - turn) % 4, Deskew = false }
                }, deadline.Token);
            await RecordAsync(id, "ocr", searchable.Ocr.Document, expected, expectedTable, caption, output, results, deadline.Token);
            string providerText = string.Join(" ", searchable.Ocr.Pages.SelectMany(static page => page.Words)
                .Select(static word => word.Text));
            await File.WriteAllTextAsync(Path.Combine(output, id + "-provider.txt"), providerText, deadline.Token);
            recognition.Add(new { Id = id, Provider = searchable.Ocr.Pages[0].Provider,
                Model = searchable.Ocr.Pages[0].Model, Language = searchable.Ocr.Pages[0].Language,
                Tokens = MeasureRecognitionTokens(string.Join(" ", expected), providerText) });
            byte[] bytes = searchable.Document.ToBytes();
            await File.WriteAllBytesAsync(Path.Combine(output, id + "-searchable.pdf"), bytes, deadline.Token);
            PdfDocumentReadResult roundTrip = PdfDocument.Load(bytes).Read(new PdfReadOptions { Profile = PdfReadProfile.Structured });
            await RecordAsync(id, "readback", roundTrip, expected, expectedTable, caption, output, results, deadline.Token);
        }
        await File.WriteAllTextAsync(Path.Combine(output, "layout-quality.json"), JsonSerializer.Serialize(new {
            Runtime = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription,
            SourceManifestSha256 = Convert.ToHexString(SHA256.HashData(await File.ReadAllBytesAsync(Path.Combine(root, "manifest.json"), deadline.Token))),
            Metric = "NFC and collapsed whitespace. Exact labelled segments, correctly ordered segment pairs, code-point CER/token WER, table rows and caption classification are separate observations. Exact table rows require one detected table, the expected row count, and matching cells at the same row and column positions.",
            Limit = "Controlled Pango/Cairo fixtures, not a population accuracy estimate. OCR uses the labelled rotation and installed language models; recognition errors remain visible.",
            RecognitionMetric = "Case-sensitive NFC token multiset precision and recall, ignoring order. This separates recognized-token evidence from canonical line/column ordering; it is not a character accuracy score.",
            ProviderRecognition = recognition,
            Cases = results
        }, new JsonSerializerOptions { WriteIndented = true }), deadline.Token);
        return 0;
    }

    private static object MeasureRecognitionTokens(string expected, string actual) {
        string[] expectedTokens = Normalize(expected).Split(' ', StringSplitOptions.RemoveEmptyEntries);
        string[] actualTokens = Normalize(actual).Split(' ', StringSplitOptions.RemoveEmptyEntries);
        var remaining = expectedTokens.GroupBy(static value => value, StringComparer.Ordinal)
            .ToDictionary(static group => group.Key, static group => group.Count(), StringComparer.Ordinal);
        int matched = 0;
        foreach (string token in actualTokens) {
            if (!remaining.TryGetValue(token, out int count) || count == 0) continue;
            remaining[token] = count - 1;
            matched++;
        }
        return new { ExpectedTokens = expectedTokens.Length, RecognizedTokens = actualTokens.Length,
            MatchedTokens = matched, Precision = actualTokens.Length == 0 ? 0D : (double)matched / actualTokens.Length,
            Recall = expectedTokens.Length == 0 ? 0D : (double)matched / expectedTokens.Length };
    }

    private static async Task<byte[]> ReadVerifiedAsync(string root, JsonElement fixture, string suffix, CancellationToken token) {
        string id = fixture.GetProperty("id").GetString()!;
        byte[] bytes = await File.ReadAllBytesAsync(Path.Combine(root, id + "-" + suffix), token);
        string hash = Convert.ToHexString(SHA256.HashData(bytes));
        if (!string.Equals(hash, fixture.GetProperty("files").GetProperty(suffix).GetString(), StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException("Fixture digest mismatch: " + id + "-" + suffix);
        return bytes;
    }

    private static async Task RecordAsync(string id, string mode, PdfDocumentReadResult document, string[] expected,
        string[][] expectedTable, string caption, string output, List<object> results, CancellationToken token) {
        string actual = Normalize(document.Text);
        int[] positions = expected.Select(segment => actual.IndexOf(Normalize(segment), StringComparison.Ordinal)).ToArray();
        int present = positions.Count(position => position >= 0), correctPairs = 0;
        for (int first = 0; first < positions.Length; first++)
            for (int second = first + 1; second < positions.Length; second++)
                if (positions[first] >= 0 && positions[second] > positions[first]) correctPairs++;
        string[][][] tables = document.Pages.SelectMany(page => page.Tables)
            .Select(table => table.Rows.Select(row => row.ToArray()).ToArray()).ToArray();
        string[] captions = document.Pages.SelectMany(page => page.Captions).Select(item => item.Text).ToArray();
        ScanTextAccuracy accuracy = ScanTextAccuracy.Measure(string.Join(" ", expected), actual);
        await File.WriteAllTextAsync(Path.Combine(output, id + "-" + mode + ".txt"), document.Text, token);
        results.Add(new {
            Id = id, Mode = mode, Accuracy = accuracy, ExactSegments = present, ExpectedSegments = expected.Length,
            CorrectReadingOrderPairs = correctPairs, ExpectedReadingOrderPairs = expected.Length * (expected.Length - 1) / 2,
            ExactCaption = captions.Any(value => Normalize(value) == Normalize(caption)), Captions = captions,
            ExactTableRows = CountExactTableRows(expectedTable, tables),
            ExpectedTableRows = expectedTable.Length, Tables = tables,
            Lines = document.Pages.SelectMany(page => page.Analysis.Lines).Select(line => new {
                line.Text, line.XStart, line.XEnd, line.BaselineY, line.RotationDegrees, line.SourceKind
            }).ToArray()
        });
        Console.WriteLine($"{id}/{mode}: CER {accuracy.CharacterErrorRate:P2}; segments {present}/{expected.Length}; order pairs {correctPairs}/{expected.Length * (expected.Length - 1) / 2}");
    }

    private static string Normalize(string text) => Regex.Replace(text.Normalize(NormalizationForm.FormC), @"\s+", " ").Trim();

    internal static int CountExactTableRows(string[][] expected, string[][][] actual) {
        if (actual.Length != 1 || actual[0].Length != expected.Length) return 0;
        int exact = 0;
        for (int row = 0; row < expected.Length; row++)
            if (expected[row].Select(Normalize).SequenceEqual(actual[0][row].Select(Normalize))) exact++;
        return exact;
    }
}
