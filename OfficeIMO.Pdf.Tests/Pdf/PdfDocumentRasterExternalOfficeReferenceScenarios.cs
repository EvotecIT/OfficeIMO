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
            byte[] actualPdf = CreateOfficeImoReferenceCandidate(scenario.ConverterId, sourcePath, actualPdfPath);
            Assert.Equal(scenario.Pages.Count, PdfCore.PdfLogicalDocument.Load(actualPdf).Pages.Count);
            WriteReviewPdfArtifact("external-reference-" + scenario.Id + ".microsoft-office", referencePdf);
            WriteReviewPdfArtifact("external-reference-" + scenario.Id + ".officeimo", actualPdf);

            if (!TryFindPdftoppm(out string rasterizerPath)) {
                if (IsRequired()) {
                    throw new InvalidOperationException("A PDF rasterizer is required for the external Microsoft Office reference gate.");
                }

                return;
            }

            var failures = new List<string>();
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
                WriteReviewArtifact(artifactStem + ".diff.png", comparison.DiffPng);

                if (!comparison.Passed) {
                    failures.Add(
                        "page " + page.Number.ToString(CultureInfo.InvariantCulture) +
                        " (" + referenceImage.Width.ToString(CultureInfo.InvariantCulture) + "x" + referenceImage.Height.ToString(CultureInfo.InvariantCulture) +
                        " vs " + actualImage.Width.ToString(CultureInfo.InvariantCulture) + "x" + actualImage.Height.ToString(CultureInfo.InvariantCulture) + ")" +
                        ": different=" + comparison.DifferentPixels.ToString(CultureInfo.InvariantCulture) +
                        "/" + comparison.TotalPixels.ToString(CultureInfo.InvariantCulture) +
                        " (" + (comparison.TotalPixels == 0 ? 0D : comparison.DifferentPixels / (double)comparison.TotalPixels).ToString("0.0000", CultureInfo.InvariantCulture) +
                        "), MAE=" + comparison.MeanAbsoluteError.ToString("0.###", CultureInfo.InvariantCulture) +
                        ", RMSE=" + comparison.RootMeanSquareError.ToString("0.###", CultureInfo.InvariantCulture) +
                        ", luminance MAE=" + comparison.MeanLuminanceError.ToString("0.###", CultureInfo.InvariantCulture));
                }
            }

            Assert.True(
                failures.Count == 0,
                scenario.Producer + " " + corpus.ProducerEnvironment.ShortVersion +
                " reference distance exceeded the pinned budget for " + scenario.Id + ": " +
                string.Join("; ", failures));
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

    private static byte[] CreateOfficeImoReferenceCandidate(string converterId, string sourcePath, string outputPath) {
        switch (converterId) {
            case "word":
                using (WordDocument document = WordDocument.Load(sourcePath, new WordLoadOptions {
                    AccessMode = DocumentAccessMode.ReadOnly
                })) {
                    document.SaveAsPdf(outputPath, new WordPdfSaveOptions {
                        IncludePageNumbers = false
                    });
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
                throw new InvalidOperationException("Unsupported external-reference converter: " + converterId + ".");
        }

        return File.ReadAllBytes(outputPath);
    }

    private static ReferenceCorpus ReadReferenceCorpus(string referenceDirectory) {
        string metadataPath = Path.Combine(referenceDirectory, "reference-corpus.json");
        ReferenceCorpus? corpus = JsonSerializer.Deserialize<ReferenceCorpus>(
            File.ReadAllText(metadataPath),
            new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
        return corpus ?? throw new InvalidOperationException("Could not deserialize the external Office reference corpus metadata.");
    }

    private static string ComputeSha256(string path) =>
        Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant();

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
    }
}
