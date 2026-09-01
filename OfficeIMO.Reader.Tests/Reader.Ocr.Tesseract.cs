using OfficeIMO.Reader;
using OfficeIMO.Reader.Ocr.Tesseract;
using System.Net;
using System.Net.Http;
using System.Security.Cryptography;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderOcrTesseractTests {
    [Fact]
    public void TesseractTsvParser_MapsLinesWordsConfidenceAndPixelGeometry() {
        const string tsv = "level\tpage_num\tblock_num\tpar_num\tline_num\tword_num\tleft\ttop\twidth\theight\tconf\ttext\n"
            + "1\t1\t0\t0\t0\t0\t0\t0\t200\t100\t-1\t\n"
            + "5\t1\t1\t1\t1\t1\t10\t20\t40\t10\t90\tInvoice\n"
            + "5\t1\t1\t1\t1\t2\t55\t20\t30\t10\t80\t1042\n"
            + "5\t1\t1\t1\t2\t1\t10\t40\t50\t12\t100\tTotal\n";

        OfficeOcrEngineResult result = TesseractTsvParser.Parse(tsv, "eng");

        Assert.Equal("Invoice 1042" + Environment.NewLine + "Total", result.Text);
        Assert.Equal(0.9D, result.Confidence!.Value, precision: 6);
        Assert.Equal(2, result.Spans.Count(span => span.Level == OfficeOcrTextSpanLevel.Line));
        Assert.Equal(3, result.Spans.Count(span => span.Level == OfficeOcrTextSpanLevel.Word));
        OfficeOcrTextSpan firstLine = Assert.Single(result.Spans, span => span.Level == OfficeOcrTextSpanLevel.Line && span.Text == "Invoice 1042");
        Assert.Equal(10D, firstLine.Region!.X);
        Assert.Equal(75D, firstLine.Region.Width);
        Assert.Equal(OfficeOcrCoordinateUnit.Pixels, firstLine.CoordinateUnit);
        Assert.Equal("tesseract-cli", result.Provider);
    }

    [Fact]
    public void TesseractOcrEngine_BuildsOptionsBeforeTsvOutputConfig() {
        var engine = new TesseractOcrEngine(new TesseractOcrEngineOptions {
            TessdataDirectory = "/models",
            EngineMode = 1,
            PageSegmentationMode = 6,
            Dpi = 300,
            AdditionalArguments = new[] { "quiet" }
        });

        IReadOnlyList<string> arguments = engine.BuildRecognitionArguments("input image.png", "result", "eng+pol");

        Assert.Equal("input image.png", arguments[0]);
        Assert.Equal("result", arguments[1]);
        Assert.Contains("eng+pol", arguments);
        Assert.Contains("/models", arguments);
        Assert.Contains("300", arguments);
        Assert.Equal("quiet", arguments[arguments.Count - 2]);
        Assert.Equal("tsv", arguments[arguments.Count - 1]);
    }

    [Fact]
    public void TesseractOcrEngine_AdvertisesRasterFormatsOnly() {
        var engine = new TesseractOcrEngine();

        Assert.Contains("image/png", engine.Capabilities.SupportedMediaTypes);
        Assert.Contains("image/jpeg", engine.Capabilities.SupportedMediaTypes);
        Assert.DoesNotContain("image/*", engine.Capabilities.SupportedMediaTypes);
        Assert.DoesNotContain("image/svg+xml", engine.Capabilities.SupportedMediaTypes);
        Assert.DoesNotContain("image/x-emf", engine.Capabilities.SupportedMediaTypes);
        Assert.DoesNotContain("image/x-wmf", engine.Capabilities.SupportedMediaTypes);
    }

    [Fact]
    public void TesseractRuntime_UsesExplicitExecutableWithoutMutatingCallerOptions() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-tesseract-runtime-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string executable = Path.Combine(directory, Environment.OSVersion.Platform == PlatformID.Win32NT ? "tesseract.exe" : "tesseract");
        File.WriteAllBytes(executable, Array.Empty<byte>());
        var options = new TesseractOcrEngineOptions { ExecutablePath = executable, Language = "eng+pol" };
        try {
            TesseractRuntimeInfo runtime = TesseractRuntime.Discover(executable);
            TesseractOcrEngine engine = TesseractOcrEngine.CreateDefault(options);

            Assert.Equal(Path.GetFullPath(executable), runtime.ExecutablePath);
            Assert.Equal(TesseractRuntimeSource.Explicit, runtime.Source);
            Assert.Equal(executable, options.ExecutablePath);
            Assert.Equal(Path.GetFullPath(executable), engine.ExecutablePath);
            Assert.Equal("eng+pol", engine.DefaultLanguage);
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public async Task TesseractLanguageData_VerifiesDownloadsAndReusesValidCache() {
        byte[] payload = Encoding.UTF8.GetBytes("pinned-language-model");
        string hash;
        using (SHA256 sha256 = SHA256.Create()) {
            hash = string.Concat(sha256.ComputeHash(payload).Select(static value => value.ToString("x2", System.Globalization.CultureInfo.InvariantCulture)));
        }
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-tessdata-" + Guid.NewGuid().ToString("N"));
        var handler = new StaticHttpHandler(payload);
        using var client = new HttpClient(handler);
        var package = new TesseractLanguageData.Package("fixture", hash, payload.LongLength, new Uri("https://example.test/fixture.traineddata"));
        var options = new TesseractLanguageDataOptions { CacheDirectory = directory, HttpClient = client };
        try {
            TesseractLanguageDataResult first = await TesseractLanguageData.EnsurePackagesAsync(new[] { package }, options, CancellationToken.None);
            TesseractLanguageDataResult second = await TesseractLanguageData.EnsurePackagesAsync(new[] { package }, options, CancellationToken.None);

            Assert.True(first.Downloaded);
            Assert.False(second.Downloaded);
            Assert.Equal(1, handler.CallCount);
            Assert.Equal(hash, first.Files[0].Sha256);
            Assert.Equal(payload, File.ReadAllBytes(first.Files[0].Path));
        } finally {
            if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public async Task TesseractLanguageData_FailsClosedOnChecksumMismatch() {
        byte[] payload = Encoding.UTF8.GetBytes("untrusted-model");
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-tessdata-" + Guid.NewGuid().ToString("N"));
        using var client = new HttpClient(new StaticHttpHandler(payload));
        var package = new TesseractLanguageData.Package("fixture", new string('0', 64), payload.LongLength, new Uri("https://example.test/fixture.traineddata"));
        try {
            await Assert.ThrowsAsync<InvalidDataException>(() => TesseractLanguageData.EnsurePackagesAsync(
                new[] { package },
                new TesseractLanguageDataOptions { CacheDirectory = directory, HttpClient = client },
                CancellationToken.None));
            Assert.False(File.Exists(Path.Combine(directory, "fixture.traineddata")));
            Assert.Empty(Directory.GetFiles(directory, "*.download-*"));
        } finally {
            if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
        }
    }

    private sealed class StaticHttpHandler : HttpMessageHandler {
        private readonly byte[] _payload;
        internal StaticHttpHandler(byte[] payload) => _payload = payload;
        internal int CallCount { get; private set; }

        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) {
            cancellationToken.ThrowIfCancellationRequested();
            CallCount++;
            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) {
                Content = new ByteArrayContent(_payload)
            });
        }
    }
}
