using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointDeckPreflightPdfFontsTests {
        private const string SourceFamily = "OfficeIMO Preflight Source";
        private const string WideFamily = "OfficeIMO Preflight Wide";

        [Fact]
        public void PdfTextMeasurementReportsTheOverflowPdfExportHasWithTheSameFonts() {
            using var stream = new MemoryStream();
            using PowerPointDocument presentation = CreateLabelDeck(stream);
            PowerPointPdfExportOptions pdfOptions = CreateWideFontOptions();
            var preflight = new PowerPointDeckPreflightOptions { DetectShapeCollisions = false, IncludeVisualSnapshotDiagnostics = false };

            // With installed fonts the label fits; the registered wide font wraps it past the frame.
            Assert.DoesNotContain(presentation.InspectPreflight(preflight).Findings, finding => finding.Code == "text-clipped");
            PowerPointDeckPreflightReport measured = presentation.InspectPreflight(preflight.WithPdfTextMeasurement(pdfOptions));
            using PdfDocument pdf = presentation.ToPdf(pdfOptions, out OfficeConversionReport report);

            Assert.Contains(measured.Findings, finding => finding.Code == "text-clipped");
            Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Code == "text-box-overflow");
        }

        [Fact]
        public void PdfTextMeasurementCopiesTheOptionsWhenSelected() {
            using var stream = new MemoryStream();
            using PowerPointDocument presentation = CreateLabelDeck(stream);
            PowerPointPdfExportOptions pdfOptions = CreateWideFontOptions();
            var preflight = new PowerPointDeckPreflightOptions { DetectShapeCollisions = false, IncludeVisualSnapshotDiagnostics = false };
            PowerPointDeckPreflightOptions selected = preflight.WithPdfTextMeasurement(pdfOptions);
            pdfOptions.PdfOptions = null;
            preflight.DetectTextOverflow = false;

            Assert.Contains(presentation.InspectPreflight(selected).Findings, finding => finding.Code == "text-clipped");
        }

        private static PowerPointDocument CreateLabelDeck(Stream stream) {
            PowerPointDocument presentation = PowerPointDocument.Create(stream);
            presentation.SlideSize.SetSize(OfficeLength.Points(400), OfficeLength.Points(225));
            PowerPointTextBox text = presentation.AddSlide().AddTextBox("Short label",
                OfficeLength.Points(20), OfficeLength.Points(20), OfficeLength.Points(200), OfficeLength.Points(40));
            text.FontSize = OfficeLength.Points(20);
            text.FontFamily = SourceFamily;
            text.TextAutoFit = PowerPointTextAutoFit.None;
            return presentation;
        }

        private static PowerPointPdfExportOptions CreateWideFontOptions() {
            var pdf = new PdfCreateOptions();
            pdf.RegisterNamedFontFamily(new PdfEmbeddedFontFamily(WideFamily,
                ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 1000, Enumerable.Range(32, 95).ToArray())));
            pdf.RegisterFontFamilySubstitution(SourceFamily, WideFamily, PdfFontFamilySubstitutionImpact.Compatible);
            return new PowerPointPdfExportOptions { PdfOptions = pdf, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic() };
        }
    }
}
