using System;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Html;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Html;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointVisualProofTests {
        [Fact]
        public void PowerPointHtmlResultsKeepSnapshotDiagnosticsScopedToEachConversion() {
            var options = new PowerPointHtmlSaveOptions {
                Profile = OfficeHtmlConversionProfile.PowerPointVisualReview
            };
            using var firstStream = new MemoryStream();
            using PowerPointPresentation firstPresentation = PowerPointPresentation.Create(firstStream);
            firstPresentation.SlideSize.SetSizePoints(160, 100);
            PowerPointSlide firstSlide = firstPresentation.AddSlide();
            firstSlide.AddRectanglePoints(500, 500, 20, 20, "Outside slide bounds");

            PowerPointToHtmlResult first = firstPresentation.ToHtmlResult(options);
            Assert.Contains(first.ImageDiagnostics, diagnostic => diagnostic.Code == "unsupported-powerpoint-shape");

            using var secondStream = new MemoryStream();
            using PowerPointPresentation secondPresentation = PowerPointPresentation.Create(secondStream);
            secondPresentation.AddSlide().AddTextBoxPoints("Clean snapshot", 20, 20, 120, 30);
            PowerPointToHtmlResult second = secondPresentation.ToHtmlResult(options);

            Assert.Empty(second.ImageDiagnostics);
            Assert.Contains(first.ImageDiagnostics, diagnostic => diagnostic.Code == "unsupported-powerpoint-shape");
        }

        [Fact]
        public void SharedSnapshotFeedsImageHtmlAndFaithfulPdfEvidence() {
            string presentationPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                PowerPointVisualProofReport report;
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(presentationPath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    slide.BackgroundColor = "FFFFFF";
                    PowerPointTextBox title = slide.AddTextBoxPoints("Shared fidelity proof", 40, 36, 480, 46);
                    title.Title = "Slide title";
                    title.SetLanguage("en-US");
                    title.Color = "172033";
                    title.FontSize = 28;
                    PowerPointTextBox body = slide.AddTextBoxPoints(
                        "One OfficeDrawing scene feeds PNG, SVG, positioned HTML, and faithful PDF.",
                        40, 110, 560, 60);
                    body.Title = "Fidelity statement";
                    body.SetLanguage("en-US");
                    body.Color = "30343B";
                    body.FontSize = 16;
                    PowerPointAutoShape accent = slide.AddRectanglePoints(40, 205, 300, 18, "Decorative accent");
                    accent.FillColor = "0098C8";
                    accent.OutlineColor = "0098C8";
                    accent.Decorative = true;
                    presentation.Save();

                    report = presentation.CreateVisualProofReport();

                    var htmlOptions = new PowerPointHtmlSaveOptions {
                        Profile = OfficeHtmlConversionProfile.PowerPointVisualReview
                    };
                    PowerPointToHtmlResult htmlResult = presentation.ToHtmlResult(htmlOptions);
                    string html = htmlResult.Value;
                    Assert.Contains("officeimo-shared-slide-snapshot", html, StringComparison.Ordinal);
                    Assert.Contains("data-officeimo-visual-owner=\"OfficeIMO.Drawing\"", html, StringComparison.Ordinal);
                    Assert.Contains("<svg", html, StringComparison.OrdinalIgnoreCase);
                    report.RecordArtifact("proof.html", "text/html", Encoding.UTF8.GetBytes(html),
                        htmlResult.ImageDiagnostics.Count);

                    var pdfOptions = new PowerPointPdfSaveOptions().UseProfile(PdfExportProfile.Faithful);
                    PdfDocumentConversionResult pdfResult = presentation.ToPdfDocumentResult(pdfOptions);
                    byte[] pdf = pdfResult.ToBytes();
                    Assert.True(pdf.Length > 100);
                    Assert.DoesNotContain(pdfResult.Warnings,
                        warning => warning.Code == "snapshot-selective-fallback");
                    report.RecordArtifact("proof.pdf", "application/pdf", pdf, pdfResult.Warnings.Count);
                }

                report.RecordArtifact("proof.pptx",
                    "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    File.ReadAllBytes(presentationPath));
                byte[] comparisonFixture = VisualBaselineTestSupport.CreateRgbPng(2, 1,
                    new byte[] { 0, 152, 200, 23, 32, 51 });
                VisualRasterComparison comparison = VisualBaselineTestSupport.CompareRasterImages(
                    comparisonFixture, comparisonFixture, channelTolerance: 0, allowedDifferentPixels: 0);
                report.RecordPerceptualComparison("perceptual-comparison-contract",
                    comparison.TotalPixels == 0 ? 0D : comparison.DifferentPixels / (double)comparison.TotalPixels,
                    0D);
                Assert.True(report.IsSuccessful, report.ToJson());
                Assert.Single(report.Slides);
                Assert.True(report.Slides[0].ShapeCount >= 3);
                Assert.Equal(64, report.Slides[0].Png.Sha256.Length);
                Assert.Equal(64, report.Slides[0].Svg.Sha256.Length);
                Assert.Equal(0, Assert.Single(report.PerceptualProofs).DifferenceRatio);
                Assert.Contains("\"sourceKind\":\"generated\"", report.ToJson(), StringComparison.Ordinal);
            } finally {
                if (File.Exists(presentationPath)) File.Delete(presentationPath);
            }
        }

        [Fact]
        public void VisualReviewSharedSnapshotHonorsTableSuppression() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream, new PowerPointCreateOptions());
            PowerPointTable table = presentation.AddSlide().AddTablePoints(1, 1, 30, 40, 180, 45);
            table.GetCell(0, 0).Text = "FILTERED REVIEW TABLE";

            string html = presentation.ToHtml(new PowerPointHtmlSaveOptions {
                Profile = OfficeHtmlConversionProfile.PowerPointVisualReview,
                IncludeTables = false,
                IncludeExtractionProof = false
            });

            Assert.Contains("officeimo-shared-slide-snapshot", html, StringComparison.Ordinal);
            Assert.DoesNotContain("FILTERED REVIEW TABLE", html, StringComparison.Ordinal);
        }

        [Fact]
        public void SanitizedPowerPointAuthoredFixtureProducesImportedProof() {
            string fixture = Path.Combine(GetRepositoryRoot(), "Assets", "PowerPointTemplates",
                "PowerPointWithTablesAndCharts.pptx");
            Assert.True(File.Exists(fixture), "Expected sanitized PowerPoint-authored fixture was not found.");

            using PowerPointPresentation presentation = PowerPointPresentation.Load(fixture, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
            PowerPointVisualProofReport report = presentation.CreateVisualProofReport("powerpoint-authored-import");

            Assert.NotEmpty(report.Slides);
            Assert.All(report.Slides, slide => {
                Assert.Equal(64, slide.Png.Sha256.Length);
                Assert.Equal(64, slide.Svg.Sha256.Length);
                Assert.True(slide.ShapeCount > 0);
                Assert.Equal(0, slide.SnapshotErrorCount);
            });
            Assert.Contains("\"sourceKind\":\"powerpoint-authored-import\"", report.ToJson(),
                StringComparison.Ordinal);
        }

        [Fact]
        public void CompatibilityReportKeepsAutomationOptInAndExternalLanesExplicit() {
            string presentationPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            string output = Path.Combine(Path.GetTempPath(), "OfficeIMO.PowerPointCompatibility", Guid.NewGuid().ToString("N"));
            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(presentationPath)) {
                    presentation.AddSlide().AddTitle("Compatibility proof");
                    presentation.Save();
                }

                PowerPointCompatibilityReport report = PowerPointCompatibilityInspector.Inspect(
                    presentationPath, new PowerPointCompatibilityOptions { OutputDirectory = output });

                Assert.Equal(PowerPointCompatibilityStatus.Passed,
                    report.Lanes.Single(lane => lane.Lane == "OpenXml").Status);
                Assert.Equal(PowerPointCompatibilityStatus.NotRun,
                    report.Lanes.Single(lane => lane.Lane == "PowerPointDesktop").Status);
                Assert.Equal(PowerPointCompatibilityStatus.NotRun,
                    report.Lanes.Single(lane => lane.Lane == "LibreOffice").Status);
                Assert.Equal(PowerPointCompatibilityStatus.NotRun,
                    report.Lanes.Single(lane => lane.Lane == "Keynote").Status);
                Assert.Equal(PowerPointCompatibilityStatus.NotRun,
                    report.Lanes.Single(lane => lane.Lane == "GoogleSlides").Status);
                report.RecordExternal("GoogleSlides", PowerPointCompatibilityStatus.Passed,
                    "Authenticated import completed without repair prompts.");
                Assert.Equal(PowerPointCompatibilityStatus.Passed,
                    report.Lanes.Single(lane => lane.Lane == "GoogleSlides").Status);
                Assert.Contains("\"lane\":\"PowerPointDesktop\"", report.ToJson(), StringComparison.Ordinal);

                PowerPointCompatibilityReport localTools = PowerPointCompatibilityInspector.Inspect(
                    presentationPath, new PowerPointCompatibilityOptions {
                        OutputDirectory = output,
                        EnableLibreOffice = true
                    });
                Assert.NotEqual(PowerPointCompatibilityStatus.Failed,
                    localTools.Lanes.Single(lane => lane.Lane == "LibreOffice").Status);
            } finally {
                if (File.Exists(presentationPath)) File.Delete(presentationPath);
                if (Directory.Exists(output)) Directory.Delete(output, recursive: true);
            }
        }

        [Fact]
        public void PowerPointDesktopReferenceRendererRunsOnlyWhenEnvironmentOptsIn() {
            if (!string.Equals(Environment.GetEnvironmentVariable("OFFICEIMO_POWERPOINT_DESKTOP_REFERENCE"),
                    "1", StringComparison.Ordinal)) return;

            string presentationPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            string output = Path.Combine(Path.GetTempPath(), "OfficeIMO.PowerPointDesktopReference", Guid.NewGuid().ToString("N"));
            try {
                byte[] snapshotPng;
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(presentationPath)) {
                    presentation.AddSlide().AddTitle("Native reference");
                    presentation.Save();
                    snapshotPng = presentation.Slides[0].ToPng();
                }
                PowerPointReferenceRenderResult result = PowerPointDesktopReferenceRenderer.TryRender(
                    presentationPath, output, enabled: true);
                Assert.True(result.IsSuccessful, result.Message);
                Assert.NotEmpty(result.ImagePaths);
                VisualRasterComparison comparison = VisualBaselineTestSupport.CompareRasterImages(
                    File.ReadAllBytes(result.ImagePaths[0]), snapshotPng,
                    channelTolerance: 16, allowedDifferentPixels: int.MaxValue);
                double differenceRatio = comparison.TotalPixels == 0
                    ? 0D : comparison.DifferentPixels / (double)comparison.TotalPixels;
                Assert.InRange(differenceRatio, 0D, 1D);
            } finally {
                if (File.Exists(presentationPath)) File.Delete(presentationPath);
                if (Directory.Exists(output)) Directory.Delete(output, recursive: true);
            }
        }

        private static string GetRepositoryRoot() {
            DirectoryInfo? directory = new DirectoryInfo(AppContext.BaseDirectory);
            while (directory != null) {
                if (File.Exists(Path.Combine(directory.FullName, "OfficeIMO.sln")) ||
                    Directory.Exists(Path.Combine(directory.FullName, "Assets", "PowerPointTemplates"))) {
                    return directory.FullName;
                }
                directory = directory.Parent;
            }
            throw new DirectoryNotFoundException("Could not locate the OfficeIMO repository root.");
        }

    }
}
