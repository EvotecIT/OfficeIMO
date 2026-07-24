using System.Globalization;
using System.Security.Cryptography;
using System.Text.Json;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfCore = OfficeIMO.Pdf;
using WordPdfSaveOptions = OfficeIMO.Word.Pdf.PdfSaveOptions;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentRasterVisualBaselineTests {
    [Theory]
    [InlineData("word-native-report")]
    [InlineData("word-business-delivery-summary")]
    [InlineData("excel-native-daily-workbook")]
    [InlineData("powerpoint-native-dense-layout")]
    public void NativeOfficeConverter_RemainsWithinPinnedMicrosoftOfficeReferenceDistance(string scenarioId) {
        string referenceDirectory = Path.Combine(GetPdfTestsProjectRoot(), "Pdf", "ReferenceBaselines");
        ReferenceCorpus corpus = ReadReferenceCorpus(referenceDirectory);
        ReferenceScenario scenario = Assert.Single(corpus.Scenarios, item => string.Equals(item.Id, scenarioId, StringComparison.Ordinal));
        string sourcePath = Path.Combine(referenceDirectory, scenario.SourcePath);
        string referencePdfPath = Path.Combine(referenceDirectory, scenario.ReferencePdfPath);

        Assert.Equal(scenario.SourceSha256, ComputeSha256(sourcePath));
        Assert.Equal(scenario.ReferencePdfSha256, ComputeSha256(referencePdfPath));

        byte[] referencePdf = File.ReadAllBytes(referencePdfPath);
        Assert.Equal(scenario.Pages.Count, PdfCore.PdfLogicalDocument.Load(referencePdf).Pages.Count);

        string workDirectory = Path.Combine(Path.GetTempPath(), "OfficeIMO.PdfExternalReference", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(workDirectory);
        string actualPdfPath = Path.Combine(workDirectory, scenario.Id + ".officeimo.pdf");

        try {
            byte[] actualPdf = CreateOfficeImoReferenceCandidate(scenario, sourcePath, actualPdfPath);
            Assert.Equal(scenario.Pages.Count, PdfCore.PdfLogicalDocument.Load(actualPdf).Pages.Count);
            WriteReviewPdfArtifact("external-reference-" + scenario.Id + ".microsoft-office", referencePdf);
            WriteReviewPdfArtifact("external-reference-" + scenario.Id + ".officeimo", actualPdf);

            if (!TryFindPdftoppm(out string rasterizerPath)) {
                WriteExternalReferenceSummary(
                    scenario,
                    rasterizerAvailable: false,
                    Array.Empty<ExternalReferencePageResult>());
                if (IsRequired()) {
                    throw new InvalidOperationException("A PDF rasterizer is required for the external Microsoft Office reference gate.");
                }

                return;
            }

            var failures = new List<string>();
            var pageResults = new List<ExternalReferencePageResult>();
            foreach (ReferencePage page in scenario.Pages) {
                string referencePrefix = Path.Combine(workDirectory, scenario.Id + ".reference.page" + page.Number.ToString(CultureInfo.InvariantCulture));
                string actualPrefix = Path.Combine(workDirectory, scenario.Id + ".actual.page" + page.Number.ToString(CultureInfo.InvariantCulture));
                RunPdftoppm(rasterizerPath, referencePdfPath, referencePrefix, workDirectory, page.Number);
                RunPdftoppm(rasterizerPath, actualPdfPath, actualPrefix, workDirectory, page.Number);

                string referenceRasterPath = referencePrefix + ".png";
                string actualRasterPath = actualPrefix + ".png";
                byte[] referenceRaster = File.ReadAllBytes(referenceRasterPath);
                byte[] actualRaster = File.ReadAllBytes(actualRasterPath);
                OfficeRasterImage referenceImage = VisualBaselineTestSupport.DecodePng(referenceRaster, "Microsoft Office reference raster is not a supported PNG file.");
                OfficeRasterImage actualImage = VisualBaselineTestSupport.DecodePng(actualRaster, "OfficeIMO candidate raster is not a supported PNG file.");
                Assert.Equal(page.RasterWidth, referenceImage.Width);
                Assert.Equal(page.RasterHeight, referenceImage.Height);
                Assert.InRange(Math.Abs(referenceImage.Width - actualImage.Width), 0, 1);
                Assert.InRange(Math.Abs(referenceImage.Height - actualImage.Height), 0, 1);

                int comparisonWidth = Math.Max(referenceImage.Width, actualImage.Width);
                int comparisonHeight = Math.Max(referenceImage.Height, actualImage.Height);
                OfficeRasterImage normalizedReference = PadRaster(referenceImage, comparisonWidth, comparisonHeight);
                OfficeRasterImage normalizedActual = PadRaster(actualImage, comparisonWidth, comparisonHeight);
                int allowedDifferentPixels = (int)Math.Ceiling(comparisonWidth * (double)comparisonHeight * scenario.MaximumDifferentPixelRatio);
                VisualRasterComparison comparison = VisualBaselineTestSupport.CompareRasterImages(
                    normalizedReference,
                    normalizedActual,
                    channelTolerance: 0,
                    allowedDifferentPixels,
                    scenario.MaximumMeanAbsoluteError,
                    scenario.MaximumRootMeanSquareError,
                    scenario.MaximumMeanLuminanceError);

                string artifactStem = "external-reference-" + scenario.Id + ".page" + page.Number.ToString(CultureInfo.InvariantCulture);
                WriteReviewArtifact(artifactStem + ".microsoft-office.png", referenceRaster);
                WriteReviewArtifact(artifactStem + ".officeimo.png", actualRaster);
                WriteReviewArtifact(artifactStem + ".overlay.png", CreateReviewOverlay(normalizedReference, normalizedActual));
                WriteReviewArtifact(artifactStem + ".diff.png", comparison.DiffPng);

                double differentPixelRatio = comparison.TotalPixels == 0
                    ? 0D
                    : comparison.DifferentPixels / (double)comparison.TotalPixels;
                pageResults.Add(new ExternalReferencePageResult(
                    page.Number,
                    referenceImage.Width,
                    referenceImage.Height,
                    actualImage.Width,
                    actualImage.Height,
                    comparison.DifferentPixels,
                    comparison.TotalPixels,
                    differentPixelRatio,
                    comparison.MeanAbsoluteError,
                    comparison.RootMeanSquareError,
                    comparison.MeanLuminanceError,
                    page.MeasuredDifferentPixelRatio,
                    page.MeasuredMeanAbsoluteError,
                    page.MeasuredRootMeanSquareError,
                    page.MeasuredMeanLuminanceError,
                    comparison.Passed));

                if (!comparison.Passed) {
                    failures.Add(
                        "page " + page.Number.ToString(CultureInfo.InvariantCulture) +
                        " (" + referenceImage.Width.ToString(CultureInfo.InvariantCulture) + "x" + referenceImage.Height.ToString(CultureInfo.InvariantCulture) +
                        " vs " + actualImage.Width.ToString(CultureInfo.InvariantCulture) + "x" + actualImage.Height.ToString(CultureInfo.InvariantCulture) + ")" +
                        ": different=" + comparison.DifferentPixels.ToString(CultureInfo.InvariantCulture) +
                        "/" + comparison.TotalPixels.ToString(CultureInfo.InvariantCulture) +
                        " (" + differentPixelRatio.ToString("0.0000", CultureInfo.InvariantCulture) +
                        "), MAE=" + comparison.MeanAbsoluteError.ToString("0.###", CultureInfo.InvariantCulture) +
                        ", RMSE=" + comparison.RootMeanSquareError.ToString("0.###", CultureInfo.InvariantCulture) +
                        ", luminance MAE=" + comparison.MeanLuminanceError.ToString("0.###", CultureInfo.InvariantCulture));
                }
            }

            WriteExternalReferenceSummary(scenario, rasterizerAvailable: true, pageResults);
            Assert.True(
                failures.Count == 0,
                scenario.Producer + " " + (scenario.ProducerVersion ?? corpus.ProducerEnvironment.ShortVersion) +
                " reference distance exceeded the pinned budget for " + scenario.Id + ": " +
                string.Join("; ", failures));
        } finally {
            TryDeleteDirectory(workDirectory);
        }
    }

    private static byte[] CreateReviewOverlay(OfficeRasterImage reference, OfficeRasterImage actual) {
        var overlay = new OfficeRasterImage(reference.Width, reference.Height, OfficeColor.White);
        for (int y = 0; y < reference.Height; y++) {
            for (int x = 0; x < reference.Width; x++) {
                OfficeColor referenceColor = reference.GetPixel(x, y);
                OfficeColor actualColor = actual.GetPixel(x, y);
                int referenceInk = 255 - Luminance(referenceColor);
                int actualInk = 255 - Luminance(actualColor);
                overlay.SetPixel(
                    x,
                    y,
                    OfficeColor.FromRgb(
                        (byte)(255 - actualInk),
                        (byte)(255 - Math.Max(referenceInk, actualInk)),
                        (byte)(255 - referenceInk)));
            }
        }

        return OfficePngWriter.Encode(overlay, OfficePngCompression.Optimal);
    }

    private static int Luminance(OfficeColor color) =>
        (int)Math.Round(
            0.299D * color.R + 0.587D * color.G + 0.114D * color.B,
            MidpointRounding.AwayFromZero);

    private static void WriteExternalReferenceSummary(
        ReferenceScenario scenario,
        bool rasterizerAvailable,
        IReadOnlyList<ExternalReferencePageResult> pages) {
        var summary = new {
            scenarioId = scenario.Id,
            converterId = scenario.ConverterId,
            producer = scenario.Producer,
            producerVersion = scenario.ProducerVersion,
            rasterizerAvailable,
            passed = rasterizerAvailable && pages.All(page => page.Passed),
            thresholds = new {
                maximumDifferentPixelRatio = scenario.MaximumDifferentPixelRatio,
                maximumMeanAbsoluteError = scenario.MaximumMeanAbsoluteError,
                maximumRootMeanSquareError = scenario.MaximumRootMeanSquareError,
                maximumMeanLuminanceError = scenario.MaximumMeanLuminanceError
            },
            overlayLegend = new {
                referenceOnly = "red",
                officeImoOnly = "blue",
                overlap = "black",
                background = "white"
            },
            pages = pages.Select(page => new {
                page = page.PageNumber,
                referenceSize = new { width = page.ReferenceWidth, height = page.ReferenceHeight },
                officeImoSize = new { width = page.OfficeImoWidth, height = page.OfficeImoHeight },
                differentPixels = page.DifferentPixels,
                totalPixels = page.TotalPixels,
                current = new {
                    differentPixelRatio = page.DifferentPixelRatio,
                    meanAbsoluteError = page.MeanAbsoluteError,
                    rootMeanSquareError = page.RootMeanSquareError,
                    meanLuminanceError = page.MeanLuminanceError
                },
                pinnedBaseline = new {
                    differentPixelRatio = page.PinnedDifferentPixelRatio,
                    meanAbsoluteError = page.PinnedMeanAbsoluteError,
                    rootMeanSquareError = page.PinnedRootMeanSquareError,
                    meanLuminanceError = page.PinnedMeanLuminanceError
                },
                delta = new {
                    differentPixelRatio = page.DifferentPixelRatio - page.PinnedDifferentPixelRatio,
                    meanAbsoluteError = page.MeanAbsoluteError - page.PinnedMeanAbsoluteError,
                    rootMeanSquareError = page.RootMeanSquareError - page.PinnedRootMeanSquareError,
                    meanLuminanceError = page.MeanLuminanceError - page.PinnedMeanLuminanceError
                },
                page.Passed
            })
        };
        WriteReviewArtifact(
            "external-reference-" + scenario.Id + ".comparison.json",
            JsonSerializer.SerializeToUtf8Bytes(summary, new JsonSerializerOptions { WriteIndented = true }));
    }

    [Fact]
    public void WordBusinessDeliverySummary_PreservesPaginationMarginsTextOrderAndTags() {
        string referenceDirectory = Path.Combine(GetPdfTestsProjectRoot(), "Pdf", "ReferenceBaselines");
        ReferenceCorpus corpus = ReadReferenceCorpus(referenceDirectory);
        ReferenceScenario scenario = Assert.Single(
            corpus.Scenarios,
            item => string.Equals(item.Id, "word-business-delivery-summary", StringComparison.Ordinal));
        string sourcePath = Path.Combine(referenceDirectory, scenario.SourcePath);
        string workDirectory = Path.Combine(Path.GetTempPath(), "OfficeIMO.PdfBusinessReference", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(workDirectory);

        try {
            byte[] actual = CreateOfficeImoReferenceCandidate(
                scenario,
                sourcePath,
                Path.Combine(workDirectory, "word-business-delivery-summary.officeimo.pdf"));
            PdfCore.PdfReadDocument readDocument = PdfCore.PdfReadDocument.Open(actual);
            Assert.Equal(9, readDocument.Pages.Count);
            Assert.Equal("en-US", readDocument.CatalogLanguage);
            Assert.True(readDocument.HasTaggedContent);
            Assert.Contains("Table", readDocument.TaggedContent!.StructureTypes);
            Assert.Contains("L", readDocument.TaggedContent.StructureTypes);

            using UglyToad.PdfPig.PdfDocument pdf = UglyToad.PdfPig.PdfDocument.Open(actual);
            UglyToad.PdfPig.Content.Page pageOne = pdf.GetPage(1);
            UglyToad.PdfPig.Content.Page pageTwo = pdf.GetPage(2);
            Assert.Contains("Planning Workbook", pageOne.Text, StringComparison.Ordinal);
            Assert.DoesNotContain("WorksheetPurpose", pageOne.Text, StringComparison.Ordinal);
            Assert.Contains("WorksheetPurpose", pageTwo.Text, StringComparison.Ordinal);
            List<UglyToad.PdfPig.Content.Word> pageTwoWords = pageTwo.GetWords().ToList();
            UglyToad.PdfPig.Content.Word firstRecommendation = pageTwoWords
                .Where(word => string.Equals(word.Text, "Recommendation:", StringComparison.Ordinal))
                .OrderByDescending(word => word.BoundingBox.Bottom)
                .First();
            UglyToad.PdfPig.Content.Word technicalStatus = pageTwoWords
                .Where(word => string.Equals(word.Text, "Technical", StringComparison.Ordinal))
                .OrderByDescending(word => word.BoundingBox.Bottom)
                .First();
            UglyToad.PdfPig.Content.Word deliveryWorksheet = pageTwoWords
                .Where(word => string.Equals(word.Text, "Delivery", StringComparison.Ordinal))
                .OrderByDescending(word => word.BoundingBox.Bottom)
                .First();
            Assert.InRange(
                firstRecommendation.BoundingBox.Bottom - technicalStatus.BoundingBox.Bottom,
                128D,
                136D);
            Assert.True(
                firstRecommendation.BoundingBox.Bottom - deliveryWorksheet.BoundingBox.Bottom >= 224D,
                "Expected Word's document-default auto line spacing and the list boundary spacing to preserve the authored page-two vertical rhythm.");
            Assert.True(
                pageTwo.Letters.Where(letter => letter.Value == "W").Max(letter => letter.StartBaseLine.Y) <= 720.5D,
                "Expected the planning table to start inside the one-inch top margin.");

            foreach (UglyToad.PdfPig.Content.Page page in pdf.GetPages()) {
                Assert.All(
                    page.Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value)),
                    letter => {
                        Assert.InRange(letter.StartBaseLine.X, 68D, 544D);
                        Assert.InRange(letter.StartBaseLine.Y, 68D, 724D);
                    });
            }
        } finally {
            TryDeleteDirectory(workDirectory);
        }
    }

    private static OfficeRasterImage PadRaster(OfficeRasterImage source, int width, int height) {
        if (source.Width == width && source.Height == height) {
            return source;
        }

        var padded = new OfficeRasterImage(width, height, OfficeColor.White);
        for (int y = 0; y < source.Height; y++) {
            for (int x = 0; x < source.Width; x++) {
                padded.SetPixel(x, y, source.GetPixel(x, y));
            }
        }

        return padded;
    }

    private static byte[] CreateOfficeImoReferenceCandidate(ReferenceScenario scenario, string sourcePath, string outputPath) {
        switch (scenario.ConverterId) {
            case "word":
                using (WordDocument document = WordDocument.Load(sourcePath, new WordLoadOptions {
                    AccessMode = DocumentAccessMode.ReadOnly
                })) {
                    var options = new WordPdfSaveOptions {
                        IncludePageNumbers = false
                    };
                    if (string.Equals(scenario.FontProfile, "officeimo-browser-compact", StringComparison.Ordinal)) {
                        options.FontFamily = "Carlito";
                        options.PdfOptions = CreateBrowserCompactReferenceOptions(scenario);
                        options.ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic();
                    }

                    document.SaveAsPdf(outputPath, options);
                }
                break;
            case "excel":
                using (ExcelDocument document = ExcelDocument.Load(sourcePath)) {
                    document.SaveAsPdf(
                        outputPath,
                        new ExcelPdfSaveOptions().UseProfile(PdfCore.PdfExportProfile.Faithful));
                }
                break;
            case "powerpoint":
                using (PowerPointPresentation presentation = PowerPointPresentation.Load(
                    sourcePath,
                    new PowerPointLoadOptions { AccessMode = DocumentAccessMode.ReadOnly })) {
                    var options = new PowerPointPdfSaveOptions {
                        ChartStyle = new OfficeChartStyle(
                            palette: new[] { OfficeColor.FromRgb(34, 126, 102) },
                            backgroundColor: OfficeColor.White,
                            titleColor: OfficeColor.FromRgb(26, 31, 43)),
                        ChartLayout = new OfficeChartLayout(maximumCategoryAxisLabels: 4)
                    };
                    File.WriteAllBytes(outputPath, presentation.ToPdf(options));
                }
                break;
            default:
                throw new InvalidOperationException("Unsupported external-reference converter: " + scenario.ConverterId + ".");
        }

        return File.ReadAllBytes(outputPath);
    }

    private static PdfCore.PdfOptions CreateBrowserCompactReferenceOptions(ReferenceScenario scenario) {
        string fontDirectory = Path.GetFullPath(Path.Combine(
            GetPdfTestsProjectRoot(),
            "..",
            "Website",
            "Apps",
            "OfficeIMO.Web.Converter",
            "Assets",
            "Fonts"));
        var assets = new SortedDictionary<string, byte[]>(StringComparer.Ordinal) {
            ["Carlito-Bold.ttf"] = File.ReadAllBytes(Path.Combine(fontDirectory, "Carlito-Bold.ttf")),
            ["Carlito-BoldItalic.ttf"] = File.ReadAllBytes(Path.Combine(fontDirectory, "Carlito-BoldItalic.ttf")),
            ["Carlito-Italic.ttf"] = File.ReadAllBytes(Path.Combine(fontDirectory, "Carlito-Italic.ttf")),
            ["Carlito-Regular.ttf"] = File.ReadAllBytes(Path.Combine(fontDirectory, "Carlito-Regular.ttf")),
            ["NotoSansArabic-Regular.ttf"] = File.ReadAllBytes(Path.Combine(fontDirectory, "NotoSansArabic-Regular.ttf")),
            ["NotoSansSymbols2-Regular.ttf"] = File.ReadAllBytes(Path.Combine(fontDirectory, "NotoSansSymbols2-Regular.ttf"))
        };
        using IncrementalHash hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        foreach (KeyValuePair<string, byte[]> asset in assets) {
            hash.AppendData(System.Text.Encoding.UTF8.GetBytes(asset.Key));
            hash.AppendData(new byte[] { 0 });
            hash.AppendData(asset.Value);
        }
        string fingerprint = BitConverter.ToString(hash.GetHashAndReset())
            .Replace("-", string.Empty)
            .ToLowerInvariant();
        Assert.Equal(scenario.FontProfileFingerprint, fingerprint);

        var family = new PdfCore.PdfEmbeddedFontFamily(
            "Carlito",
            assets["Carlito-Regular.ttf"],
            assets["Carlito-Bold.ttf"],
            assets["Carlito-Italic.ttf"],
            assets["Carlito-BoldItalic.ttf"]);
        var options = new PdfCore.PdfOptions {
            DefaultFont = PdfCore.PdfStandardFont.Helvetica,
            HeaderFont = PdfCore.PdfStandardFont.Helvetica,
            FooterFont = PdfCore.PdfStandardFont.Helvetica,
            TaggedStructureMode = PdfCore.PdfTaggedStructureMode.CatalogMarkers,
            TextShapingMode = PdfCore.PdfTextShapingMode.LatinLigatures
        };
        options.RegisterFontFamily(PdfCore.PdfStandardFont.Helvetica, family);
        options.RegisterNamedFontFamily(family);
        options.RegisterEmbeddedFontFallbacks(
            new PdfCore.PdfEmbeddedFontFallbackSet(new[] {
                new PdfCore.PdfEmbeddedFontFallbackCandidate("Noto Sans Arabic", assets["NotoSansArabic-Regular.ttf"]),
                new PdfCore.PdfEmbeddedFontFallbackCandidate("Noto Sans Symbols 2", assets["NotoSansSymbols2-Regular.ttf"])
            }));
        return options;
    }

    private static ReferenceCorpus ReadReferenceCorpus(string referenceDirectory) {
        string metadataPath = Path.Combine(referenceDirectory, "reference-corpus.json");
        ReferenceCorpus? corpus = JsonSerializer.Deserialize<ReferenceCorpus>(
            File.ReadAllText(metadataPath),
            new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
        return corpus ?? throw new InvalidOperationException("Could not deserialize the external Office reference corpus metadata.");
    }

    private static string ComputeSha256(string path) {
        using SHA256 sha256 = SHA256.Create();
        using FileStream stream = File.OpenRead(path);
        return BitConverter.ToString(sha256.ComputeHash(stream))
            .Replace("-", string.Empty)
            .ToLowerInvariant();
    }

    private sealed class ReferenceCorpus {
        public ReferenceProducerEnvironment ProducerEnvironment { get; set; } = new();

        public List<ReferenceScenario> Scenarios { get; set; } = new();
    }

    private sealed class ReferenceProducerEnvironment {
        public string ShortVersion { get; set; } = string.Empty;
    }

    private sealed class ReferenceScenario {
        public string Id { get; set; } = string.Empty;

        public string ConverterId { get; set; } = string.Empty;

        public string Producer { get; set; } = string.Empty;

        public string? ProducerVersion { get; set; }

        public string? FontProfile { get; set; }

        public string? FontProfileFingerprint { get; set; }

        public string SourcePath { get; set; } = string.Empty;

        public string SourceSha256 { get; set; } = string.Empty;

        public string ReferencePdfPath { get; set; } = string.Empty;

        public string ReferencePdfSha256 { get; set; } = string.Empty;

        public List<ReferencePage> Pages { get; set; } = new();

        public double MaximumDifferentPixelRatio { get; set; }

        public double MaximumMeanAbsoluteError { get; set; }

        public double MaximumRootMeanSquareError { get; set; }

        public double MaximumMeanLuminanceError { get; set; }
    }

    private sealed class ReferencePage {
        public int Number { get; set; }

        public int RasterWidth { get; set; }

        public int RasterHeight { get; set; }

        public double MeasuredDifferentPixelRatio { get; set; }

        public double MeasuredMeanAbsoluteError { get; set; }

        public double MeasuredRootMeanSquareError { get; set; }

        public double MeasuredMeanLuminanceError { get; set; }
    }

    private sealed record ExternalReferencePageResult(
        int PageNumber,
        int ReferenceWidth,
        int ReferenceHeight,
        int OfficeImoWidth,
        int OfficeImoHeight,
        int DifferentPixels,
        int TotalPixels,
        double DifferentPixelRatio,
        double MeanAbsoluteError,
        double RootMeanSquareError,
        double MeanLuminanceError,
        double PinnedDifferentPixelRatio,
        double PinnedMeanAbsoluteError,
        double PinnedRootMeanSquareError,
        double PinnedMeanLuminanceError,
        bool Passed);
}
