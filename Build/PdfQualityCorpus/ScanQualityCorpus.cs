using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using OfficeIMO.Drawing;
using OfficeIMO.Ocr;
using OfficeIMO.Ocr.Tesseract;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;

namespace OfficeIMO.PdfQualityCorpus;

// Opt-in integration evidence: requires an installed Tesseract executable and eng/osd data.
internal static class ScanQualityCorpus {
    internal static async Task<int> RunAsync(string[] args) {
        if (args.Length != 3) throw new ArgumentException("Usage: scan <ScanQuality fixture directory> <output directory>");
        string root = Path.GetFullPath(args[1]), output = Path.GetFullPath(args[2]);
        Directory.CreateDirectory(output);
        using var deadline = new CancellationTokenSource(TimeSpan.FromMinutes(10));
        var engine = new TesseractOcrEngine(new TesseractOcrEngineOptions {
            Dpi = 300, Language = "eng", PageSegmentationMode = 3, Timeout = TimeSpan.FromSeconds(30),
            TemporaryDirectory = Path.Combine(output, "temporary")
        });
        var cases = new List<ScanCase>();
        foreach (string name in new[] { "phototest", "eurotext" }) {
            string truthPath = Path.Combine(root, name == "phototest" ? "phototest.gold.txt" : "eurotext.txt");
            string truth = await File.ReadAllTextAsync(truthPath, deadline.Token);
            foreach (string variant in new[] { "", "-skew", "-shadow-skew", "-clockwise90", "-upside-down" }) {
                string id = name + variant;
                byte[] input = await File.ReadAllBytesAsync(Path.Combine(root, id + ".pdf"), deadline.Token);
                var item = new ScanCase { Id = id, InputSha256 = Convert.ToHexString(SHA256.HashData(input)),
                    TruthSha256 = Convert.ToHexString(SHA256.HashData(await File.ReadAllBytesAsync(truthPath, deadline.Token))) };
                item.Before = await RunCaseAsync(engine, input, truth, output, id + "-before", false, deadline.Token);
                item.After = await RunCaseAsync(engine, input, truth, output, id + "-after", true, deadline.Token);
                cases.Add(item);
                Console.WriteLine($"{id}: CER {item.Before.ProviderAccuracy.CharacterErrorRate:P2} -> {item.After.ProviderAccuracy.CharacterErrorRate:P2}; " +
                    $"WER {item.Before.ProviderAccuracy.WordErrorRate:P2} -> {item.After.ProviderAccuracy.WordErrorRate:P2}; preserved={item.After.OriginalAppearancePreserved}");
            }
        }
        var report = new { Provider = engine.Id, ProviderVersion = await engine.GetVersionAsync(deadline.Token),
            Runtime = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription,
            OperatingSystem = System.Runtime.InteropServices.RuntimeInformation.OSDescription,
            Language = "eng", Dpi = 300, MinimumOrientationConfidence = 0.75,
            Metric = "NFC text with collapsed whitespace; case and punctuation retained; Levenshtein code-point CER and token WER. Rates can exceed 1.",
            Limit = "Two upstream labelled scans with deterministic degradations; English model only, including multilingual eurotext. Buffer accounting excludes provider process and PDF encoding.",
            Cases = cases };
        await File.WriteAllTextAsync(Path.Combine(output, "scan-quality.json"), JsonSerializer.Serialize(report,
            new JsonSerializerOptions { WriteIndented = true }), deadline.Token);
        return cases.All(c => c.Before.OriginalAppearancePreserved && c.After.OriginalAppearancePreserved) ? 0 : 1;
    }

    private static async Task<ScanRun> RunCaseAsync(IOcrEngine engine, byte[] input, string truth, string output,
        string id, bool cleanup, CancellationToken token) {
        OcrOrientationResult? orientation = null;
        string providerText = "";
        var recording = new DelegateOcrEngine(engine.Id + "-corpus", async (request, cancellation) => {
            OcrResult result = await engine.RecognizeAsync(request, cancellation).ConfigureAwait(false);
            if (request.Operation == OcrOperation.DetectOrientation) orientation = result.Orientation;
            else {
                await File.WriteAllBytesAsync(Path.Combine(output, id + ".png"), request.Payload, cancellation);
                providerText = string.Join(" ", result.Spans.Where(span => span.Level == OcrTextSpanLevel.Word).Select(span => span.Text));
            }
            return result;
        }, engine.Capabilities);
        PdfDocument document = PdfDocument.Load(input);
        PdfSearchableOcrReview review = await document.PrepareSearchableOcrAsync(recording, new PdfOcrMergeOptions {
            Dpi = 300, MinimumConfidence = 0, DetectOrientation = cleanup, ProviderTimeout = TimeSpan.FromSeconds(35),
            ScanProcessing = cleanup ? new OfficeScanProcessingOptions() : null
        }, token);
        PdfSearchableOcrResult written = review.ApplyAll(token);
        byte[] searchable = written.Document.ToBytes();
        await File.WriteAllBytesAsync(Path.Combine(output, id + ".pdf"), searchable, token);
        await File.WriteAllTextAsync(Path.Combine(output, id + "-provider.txt"), providerText, token);
        await File.WriteAllTextAsync(Path.Combine(output, id + "-extracted.txt"), review.Ocr.Text, token);
        string roundTrip = PdfDocument.Load(searchable).Read(new PdfReadOptions { Profile = PdfReadProfile.Structured }).Text;
        await File.WriteAllTextAsync(Path.Combine(output, id + "-roundtrip.txt"), roundTrip, token);
        byte[] Render(byte[] bytes) => PdfDocument.Load(bytes).Render.Pages(PdfPageSelection.From(new PdfPageRange(1, 1)),
            new PdfPageRenderOptions { Format = PdfPageRenderFormat.Png, Dpi = 72, ContinueOnError = false })[0].Bytes
            ?? throw new InvalidOperationException("The preserved page did not render.");
        byte[] sourceRaster = Render(input);
        byte[] writtenRaster = Render(searchable);
        return new ScanRun {
            ProviderAccuracy = ScanTextAccuracy.Measure(truth, providerText),
            ReconstructedAccuracy = ScanTextAccuracy.Measure(truth, review.Ocr.Text),
            RoundTripAccuracy = ScanTextAccuracy.Measure(truth, roundTrip),
            OriginalAppearancePreserved = sourceRaster.AsSpan().SequenceEqual(writtenRaster) && input.AsSpan().SequenceEqual(document.ToBytes()),
            Orientation = orientation, Processing = review.Ocr.Pages[0].ScanProcessing,
            Diagnostics = review.Ocr.Pages[0].Diagnostics.ToArray(), Words = review.Ocr.AcceptedWordCount
        };
    }

    private sealed class ScanCase {
        public string Id { get; init; } = "";
        public string InputSha256 { get; init; } = "";
        public string TruthSha256 { get; init; } = "";
        public ScanRun Before { get; set; } = new();
        public ScanRun After { get; set; } = new();
    }
    private sealed class ScanRun {
        public ScanTextAccuracy ProviderAccuracy { get; init; } = new();
        public ScanTextAccuracy ReconstructedAccuracy { get; init; } = new();
        public ScanTextAccuracy RoundTripAccuracy { get; init; } = new();
        public bool OriginalAppearancePreserved { get; init; }
        public int Words { get; init; }
        public OcrOrientationResult? Orientation { get; init; }
        public OfficeScanProcessingReport? Processing { get; init; }
        public string[] Diagnostics { get; init; } = Array.Empty<string>();
    }
}
