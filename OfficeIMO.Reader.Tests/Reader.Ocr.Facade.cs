using OfficeIMO.Pdf;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Ocr;
using OfficeIMO.Reader.Ocr.Tesseract;
using OfficeIMO.Tests.Pdf;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderOcrFacadeTests {
    [Theory]
    [InlineData(OfficeOcrLanguage.English, "eng")]
    [InlineData(OfficeOcrLanguage.Polish, "pol")]
    [InlineData(OfficeOcrLanguage.English | OfficeOcrLanguage.Polish, "eng+pol")]
    [InlineData(OfficeOcrLanguage.French | OfficeOcrLanguage.German | OfficeOcrLanguage.Spanish, "fra+deu+spa")]
    [InlineData(OfficeOcrLanguage.Arabic | OfficeOcrLanguage.Hebrew | OfficeOcrLanguage.Hindi, "ara+heb+hin")]
    [InlineData(OfficeOcrLanguage.ChineseSimplified | OfficeOcrLanguage.Japanese | OfficeOcrLanguage.Korean, "chi_sim+jpn+kor")]
    public void Languages_MapDiscoverableValuesToStableProviderExpressions(OfficeOcrLanguage languages, string expected) {
        Assert.Equal(expected, languages.ToTesseractExpression());
    }

    [Fact]
    public void Languages_RejectEmptyAndUndefinedSelections() {
        Assert.Throws<ArgumentOutOfRangeException>(() => ((OfficeOcrLanguage) 0).ToTesseractExpression());
        Assert.Throws<ArgumentOutOfRangeException>(() => ((OfficeOcrLanguage) (1UL << 40)).ToTesseractExpression());
    }

    [Fact]
    public void Languages_ExposeOneTypedEntryForEveryProvisionedFacadeModel() {
        Assert.Equal(28, OfficeOcrLanguages.Supported.Count);
        Assert.Equal(OfficeOcrLanguages.Supported.Count, OfficeOcrLanguages.Supported.Distinct().Count());
        Assert.All(OfficeOcrLanguages.Supported, language => {
            string code = language.ToTesseractExpression();
            Assert.Contains(code, TesseractLanguageData.SupportedLanguages);
        });
    }

    [Fact]
    public void Languages_PreserveLegacyConfigurationAndRejectAmbiguousOverrides() {
        Assert.Equal("eng", OfficeOcr.ResolveLanguageExpression(new OfficeOcrOptions()));
        Assert.Equal("eng+pol", OfficeOcr.ResolveLanguageExpression(new OfficeOcrOptions {
            Languages = OfficeOcrLanguage.English | OfficeOcrLanguage.Polish
        }));
        Assert.Equal("pol", OfficeOcr.ResolveLanguageExpression(new OfficeOcrOptions {
            Tesseract = new TesseractOcrEngineOptions { Language = "pol" }
        }));
        Assert.Equal("deu", OfficeOcr.ResolveLanguageExpression(new OfficeOcrOptions {
            CustomLanguageExpression = "deu"
        }));

        Assert.Throws<ArgumentException>(() => OfficeOcr.ResolveLanguageExpression(new OfficeOcrOptions {
            Languages = OfficeOcrLanguage.English | OfficeOcrLanguage.Polish,
            Tesseract = new TesseractOcrEngineOptions { Language = "pol" }
        }));
        Assert.Throws<ArgumentException>(() => OfficeOcr.ResolveLanguageExpression(new OfficeOcrOptions {
            CustomLanguageExpression = "deu",
            Tesseract = new TesseractOcrEngineOptions { Language = "pol" }
        }));
        Assert.Throws<ArgumentException>(() => OfficeOcr.ResolveLanguageExpression(new OfficeOcrOptions {
            Languages = OfficeOcrLanguage.Polish,
            CustomLanguageExpression = "deu"
        }));
    }

    [Theory]
    [InlineData(null, false)]
    [InlineData(0, true)]
    [InlineData(1, true)]
    [InlineData(2, false)]
    [InlineData(3, false)]
    [InlineData(11, false)]
    [InlineData(12, true)]
    [InlineData(13, false)]
    public void Session_RequiresOrientationDataOnlyForOsdSegmentationModes(int? pageSegmentationMode, bool expectsOsd) {
        string[] required = OfficeOcr.ResolveRequiredLanguageData("eng+pol", pageSegmentationMode);

        Assert.Equal(expectsOsd, required.Contains("osd", StringComparer.Ordinal));
        Assert.Contains("eng", required);
        Assert.Contains("pol", required);
        Assert.Equal(required.Length, required.Distinct(StringComparer.Ordinal).Count());
    }

    [Fact]
    public async Task ReadTextAsync_RejectsNullNestedTesseractOptionsBeforeReadingTheFile() {
        var options = new OfficeOcrOptions { Tesseract = null! };

        ArgumentException exception = await Assert.ThrowsAsync<ArgumentException>(() =>
            OfficeOcr.ReadTextAsync("missing-image.png", options));

        Assert.Equal("options", exception.ParamName);
        Assert.Contains("Tesseract options", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public async Task MakePdfSearchableAsync_RejectsExistingOutputBeforeRuntimeDiscovery() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-ocr-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string outputPath = Path.Combine(directory, "existing.pdf");
        try {
            File.WriteAllBytes(outputPath, new byte[] { 1, 2, 3 });
            var options = new OfficeOcrOptions {
                Tesseract = new TesseractOcrEngineOptions { ExecutablePath = "missing-tesseract-fixture" }
            };

            IOException exception = await Assert.ThrowsAsync<IOException>(() =>
                OfficeOcr.MakePdfSearchableAsync(Path.Combine(directory, "missing.pdf"), outputPath, options));

            Assert.Contains(Path.GetFullPath(outputPath), exception.Message, StringComparison.Ordinal);
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void ResolvePdfPaths_SnapshotsRelativePathsAgainstTheCurrentDirectory() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-ocr-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        lock (ReaderCurrentDirectoryLock.Gate) {
            string originalDirectory = Environment.CurrentDirectory;
            try {
                Environment.CurrentDirectory = directory;

                (string inputPath, string outputPath) = OfficeOcr.ResolvePdfPaths("input.pdf", "output.pdf");

                Assert.Equal(Path.Combine(directory, "input.pdf"), inputPath);
                Assert.Equal(Path.Combine(directory, "output.pdf"), outputPath);
            } finally {
                Environment.CurrentDirectory = originalDirectory;
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public async Task Session_RecognizesImageWithStableSourceEvidenceAndConfiguredLanguage() {
        OfficeOcrEngineRequest? captured = null;
        var engine = new DelegateOfficeOcrEngine("fixture", (request, _) => {
            captured = request;
            return new ValueTask<OfficeOcrEngineResult>(new OfficeOcrEngineResult {
                Text = "Recognized",
                Provider = "fixture"
            });
        });
        OfficeOcrSession session = CreateSession(engine, "eng+pol", new PdfOcrMergeOptions());

        OfficeOcrEngineResult result = await session.RecognizeImageAsync(new byte[] { 1, 2, 3 }, "image/png", "scan.png");

        Assert.Equal("Recognized", result.Text);
        Assert.NotNull(captured);
        Assert.Equal("eng+pol", captured!.Language);
        Assert.Equal("image/png", captured.Asset.MediaType);
        Assert.Equal(".png", captured.Asset.Extension);
        Assert.Equal("scan.png", captured.Asset.FileName);
        Assert.Equal(64, captured.Source.SourceHash!.Length);
        Assert.Equal(captured.Source.SourceHash, captured.Asset.PayloadHash);
        Assert.NotSame(captured.Payload, captured.Asset.PayloadBytes);
    }

    [Fact]
    public async Task Session_SnapshotsPdfPolicyAndUsesEngineNeutralWordGeometry() {
        var policy = new PdfOcrMergeOptions {
            MinimumConfidence = 0.5D,
            DetectAlignedTables = false
        };
        var engine = new DelegateOfficeOcrEngine(
            "fixture",
            (_, _) => new ValueTask<OfficeOcrEngineResult>(new OfficeOcrEngineResult {
                Text = "Searchable",
                Confidence = 0.8D,
                Provider = "fixture",
                Language = "eng",
                Spans = new[] {
                    new OfficeOcrTextSpan {
                        Sequence = 0,
                        Level = OfficeOcrTextSpanLevel.Word,
                        Text = "Searchable",
                        Confidence = 0.8D,
                        CoordinateUnit = OfficeOcrCoordinateUnit.Normalized,
                        Region = new OfficeDocumentRegion { X = 0.1D, Y = 0.1D, Width = 0.3D, Height = 0.05D }
                    }
                }
            }),
            new OfficeOcrEngineCapabilities {
                SupportedMediaTypes = new[] { "image/png" },
                SupportsWordSpans = true,
                SupportsConfidence = true
            });
        OfficeOcrSession session = CreateSession(engine, "eng", policy);
        policy.MinimumConfidence = 0.9D;
        byte[] source = PdfDocument.Create()
            .Image(PdfPngTestImages.CreateRgbPng(245, 245, 245), 220, 120)
            .ToBytes();

        PdfSearchableOcrResult result = await session.MakePdfSearchableAsync(PdfDocument.Load(source));

        Assert.Equal(1, result.AddedWordCount);
        Assert.Contains("Searchable", PdfReadDocument.Open(result.Document.ToBytes()).ExtractText(), StringComparison.Ordinal);
    }

    private static OfficeOcrSession CreateSession(
        IOfficeOcrEngine engine,
        string language,
        PdfOcrMergeOptions options) {
        var runtime = new TesseractRuntimeInfo("fixture-tesseract", null, TesseractRuntimeSource.Explicit);
        var evidence = new OfficeOcrRuntimeEvidence(runtime, "fixture-version", new[] { "eng", "pol" }, null);
        return new OfficeOcrSession(engine, language, options, evidence);
    }
}
