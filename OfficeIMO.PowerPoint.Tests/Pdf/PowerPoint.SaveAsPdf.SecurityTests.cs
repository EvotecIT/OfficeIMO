using System;
using System.IO;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests.Pdf;

public class PowerPointSaveAsPdfSecurityTests {
    [Fact]
    public void SaveAsPdf_PowerPointPresentation_StopsAtConfiguredGroupDepth() {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".pptx");
        try {
            using PowerPointPresentation presentation = PowerPointPresentation.Create(path);
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointTextBox textBox = slide.AddTextBoxPoints("Nested text", 72, 72, 144, 36);
            PowerPointTextBox secondTextBox = slide.AddTextBoxPoints("Nested sibling", 72, 120, 144, 36);
            slide.GroupShapes(new PowerPointShape[] { textBox, secondTextBox }, "Outer");

            var options = new PowerPointPdfSaveOptions {
                MaxGroupShapeDepth = 0
            };

            PdfCore.PdfDocumentConversionResult result = presentation.ToPdfDocumentResult(options);

            PdfCore.PdfConversionWarning warning = Assert.Single(result.Warnings, item => item.Code == "group-depth-limit");
            Assert.Equal("Slide 1", warning.Source);
            Assert.Contains("MaxGroupShapeDepth", warning.Message, StringComparison.Ordinal);
        } finally {
            if (File.Exists(path)) {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public void SaveAsPdf_PowerPointHandouts_HonorZeroGroupDepth() {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".pptx");
        try {
            using PowerPointPresentation presentation = PowerPointPresentation.Create(path);
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointTextBox textBox = slide.AddTextBoxPoints("Nested text", 72, 72, 144, 36);
            PowerPointTextBox secondTextBox = slide.AddTextBoxPoints("Nested sibling", 72, 120, 144, 36);
            slide.GroupShapes(new PowerPointShape[] { textBox, secondTextBox }, "Outer");

            PdfCore.PdfDocumentConversionResult result = presentation.ToPdfDocumentResult(new PowerPointPdfSaveOptions {
                PageLayout = PowerPointPdfPageLayout.Handouts,
                HandoutSlidesPerPage = 1,
                MaxGroupShapeDepth = 0
            });

            Assert.Contains(result.Warnings, warning =>
                warning.Message.Contains(nameof(PowerPointPdfSaveOptions.MaxGroupShapeDepth), StringComparison.Ordinal));
        } finally {
            if (File.Exists(path)) {
                File.Delete(path);
            }
        }
    }
}
