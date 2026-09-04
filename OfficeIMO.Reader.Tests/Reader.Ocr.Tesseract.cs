using OfficeIMO.Ocr;
using OfficeIMO.Reader;
using OfficeIMO.Ocr.Tesseract;
using System.Net;
using System.Net.Http;
using System.Security.Cryptography;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

[Collection(TesseractEnvironmentCollection.Name)]
public sealed class ReaderOcrTesseractTests {
    [Fact]
    public void TesseractTsvParser_MapsLinesWordsConfidenceAndPixelGeometry() {
        const string tsv = "level\tpage_num\tblock_num\tpar_num\tline_num\tword_num\tleft\ttop\twidth\theight\tconf\ttext\n"
            + "1\t1\t0\t0\t0\t0\t0\t0\t200\t100\t-1\t\n"
            + "5\t1\t1\t1\t1\t1\t10\t20\t40\t10\t90\tInvoice\n"
            + "5\t1\t1\t1\t1\t2\t55\t20\t30\t10\t80\t1042\n"
            + "5\t1\t1\t1\t2\t1\t10\t40\t50\t12\t100\tTotal\n";

        OcrResult result = TesseractTsvParser.Parse(tsv, "eng");

        Assert.Equal("Invoice 1042" + Environment.NewLine + "Total", result.Text);
        Assert.Equal(0.9D, result.Confidence!.Value, precision: 6);
        Assert.Equal(2, result.Spans.Count(span => span.Level == OcrTextSpanLevel.Line));
        Assert.Equal(3, result.Spans.Count(span => span.Level == OcrTextSpanLevel.Word));
        OcrTextSpan firstLine = Assert.Single(result.Spans, span => span.Level == OcrTextSpanLevel.Line && span.Text == "Invoice 1042");
        Assert.Equal(10D, firstLine.Region!.X);
        Assert.Equal(75D, firstLine.Region.Width);
        Assert.Equal(OcrCoordinateUnit.Pixels, firstLine.CoordinateUnit);
        Assert.Equal("1:1", firstLine.BlockId);
        Assert.Equal("1:1:1", firstLine.ParagraphId);
        Assert.Equal("1:1:1:1", firstLine.LineId);
        Assert.All(
            result.Spans.Where(span => span.Text == "Invoice" || span.Text == "1042"),
            span => Assert.Equal(firstLine.LineId, span.LineId));
        Assert.Equal("tesseract-cli", result.Provider);
    }

    [Fact]
    public void TesseractOcrEngine_EnablesTsvWithoutAConfigFileDependency() {
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
        Assert.Equal("quiet", arguments[arguments.Count - 3]);
        Assert.Equal("-c", arguments[arguments.Count - 2]);
        Assert.Equal("tessedit_create_tsv=1", arguments[arguments.Count - 1]);
        Assert.DoesNotContain("tsv", arguments);
    }

    [Fact]
    public void TesseractOcrEngine_AdvertisesRasterFormatsOnly() {
        var engine = new TesseractOcrEngine();

        Assert.Equal("tesseract", engine.ExecutablePath);
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
#if NET8_0_OR_GREATER
        if (!OperatingSystem.IsWindows()) {
            File.SetUnixFileMode(executable, UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute);
        }
#endif
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
    public void TesseractRuntime_DefaultOptionsPreserveEnvironmentDiscovery() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-tesseract-default-runtime-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string executable = Path.Combine(directory, Environment.OSVersion.Platform == PlatformID.Win32NT ? "tesseract.exe" : "tesseract");
        File.WriteAllBytes(executable, Array.Empty<byte>());
#if NET8_0_OR_GREATER
        if (!OperatingSystem.IsWindows()) {
            File.SetUnixFileMode(executable, UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute);
        }
#endif
        string? previous = Environment.GetEnvironmentVariable("OFFICEIMO_TESSERACT_PATH");
        var options = new TesseractOcrEngineOptions();
        try {
            Environment.SetEnvironmentVariable("OFFICEIMO_TESSERACT_PATH", executable);

            TesseractOcrEngine engine = TesseractOcrEngine.CreateDefault(options);

            Assert.Null(options.ExecutablePath);
            Assert.Equal(Path.GetFullPath(executable), engine.ExecutablePath);
        } finally {
            Environment.SetEnvironmentVariable("OFFICEIMO_TESSERACT_PATH", previous);
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void TesseractRuntime_PathSearchSkipsNonExecutableUnixFiles() {
#if NET8_0_OR_GREATER
        if (OperatingSystem.IsWindows()) return;
        string root = Path.Combine(Path.GetTempPath(), "officeimo-tesseract-path-" + Guid.NewGuid().ToString("N"));
        string blockedDirectory = Path.Combine(root, "blocked");
        string executableDirectory = Path.Combine(root, "executable");
        Directory.CreateDirectory(blockedDirectory);
        Directory.CreateDirectory(executableDirectory);
        string blocked = Path.Combine(blockedDirectory, "tesseract");
        string executable = Path.Combine(executableDirectory, "tesseract");
        File.WriteAllBytes(blocked, Array.Empty<byte>());
        File.WriteAllBytes(executable, Array.Empty<byte>());
        File.SetUnixFileMode(blocked, UnixFileMode.UserRead | UnixFileMode.UserWrite);
        File.SetUnixFileMode(executable, UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute);
        try {
            string searchPath = blockedDirectory + Path.PathSeparator + executableDirectory;

            Assert.True(TesseractRuntime.TryFindOnPath("tesseract", searchPath, out string? discovered));
            Assert.Equal(executable, discovered);
        } finally {
            Directory.Delete(root, recursive: true);
        }
#endif
    }

    [Fact]
    public void TesseractRuntime_PrefersNestedTessdataDirectoryWhenPrefixIsAnInstallationRoot() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-tesseract-prefix-" + Guid.NewGuid().ToString("N"));
        string bin = Path.Combine(root, "bin");
        string tessdata = Path.Combine(root, "tessdata");
        Directory.CreateDirectory(bin);
        Directory.CreateDirectory(tessdata);
        File.WriteAllBytes(Path.Combine(tessdata, "eng.traineddata"), Array.Empty<byte>());
        string executable = Path.Combine(bin, Environment.OSVersion.Platform == PlatformID.Win32NT ? "tesseract.exe" : "tesseract");
        File.WriteAllBytes(executable, Array.Empty<byte>());
#if NET8_0_OR_GREATER
        if (!OperatingSystem.IsWindows()) {
            File.SetUnixFileMode(executable, UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute);
        }
#endif
        string? previous = Environment.GetEnvironmentVariable("TESSDATA_PREFIX");
        try {
            Environment.SetEnvironmentVariable("TESSDATA_PREFIX", root);

            TesseractRuntimeInfo runtime = TesseractRuntime.Discover(executable);

            Assert.Equal(tessdata, runtime.TessdataDirectory);
        } finally {
            Environment.SetEnvironmentVariable("TESSDATA_PREFIX", previous);
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void TesseractRuntime_DoesNotFallBackWhenACallerSuppliesAMissingExecutableName() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-tesseract-explicit-name-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string executable = Path.Combine(root, Environment.OSVersion.Platform == PlatformID.Win32NT ? "tesseract.exe" : "tesseract");
        File.WriteAllBytes(executable, Array.Empty<byte>());
#if NET8_0_OR_GREATER
        if (!OperatingSystem.IsWindows()) {
            File.SetUnixFileMode(executable, UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute);
        }
#endif
        string? previous = Environment.GetEnvironmentVariable("OFFICEIMO_TESSERACT_PATH");
        try {
            Environment.SetEnvironmentVariable("OFFICEIMO_TESSERACT_PATH", executable);

            Assert.False(TesseractRuntime.TryDiscover("missing-tesseract-command", out TesseractRuntimeInfo? runtime));
            Assert.Null(runtime);
            Assert.Throws<FileNotFoundException>(() => TesseractRuntime.Discover("missing-tesseract-command"));
            Assert.Throws<FileNotFoundException>(() => TesseractOcrEngine.CreateDefault(new TesseractOcrEngineOptions {
                ExecutablePath = "missing-tesseract-command"
            }));
        } finally {
            Environment.SetEnvironmentVariable("OFFICEIMO_TESSERACT_PATH", previous);
            Directory.Delete(root, recursive: true);
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
            Task<TesseractLanguageDataResult> firstTask = TesseractLanguageData.EnsurePackagesAsync(new[] { package }, options, CancellationToken.None);
            Task<TesseractLanguageDataResult> secondTask = TesseractLanguageData.EnsurePackagesAsync(new[] { package }, options, CancellationToken.None);
            TesseractLanguageDataResult[] results = await Task.WhenAll(firstTask, secondTask);

            Assert.Single(results, static result => result.Downloaded);
            Assert.Single(results, static result => !result.Downloaded);
            Assert.Equal(1, handler.CallCount);
            Assert.All(results, result => Assert.Equal(hash, result.Files[0].Sha256));
            Assert.Equal(payload, File.ReadAllBytes(results[0].Files[0].Path));
            Assert.Empty(Directory.GetFiles(directory, "*.download-*"));
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
        private int _callCount;
        internal int CallCount => Volatile.Read(ref _callCount);

        protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) {
            cancellationToken.ThrowIfCancellationRequested();
            Interlocked.Increment(ref _callCount);
            await Task.Yield();
            return new HttpResponseMessage(HttpStatusCode.OK) {
                Content = new ByteArrayContent(_payload)
            };
        }
    }
}

[CollectionDefinition(Name, DisableParallelization = true)]
public sealed class TesseractEnvironmentCollection {
    public const string Name = "Tesseract environment";
}
