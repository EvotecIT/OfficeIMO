using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentComplianceAssessmentTests {

    [Fact]
    public void ImageAlternativeTextRejectsWhitespace() {
        Assert.Throws<ArgumentException>(() => new PdfImageStyle {
            AlternativeText = " "
        });

        Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().Image(CreateMinimalRgbPng(), 24, 24, alternativeText: " "));

        Assert.Throws<ArgumentException>(() => new PdfHeaderFooterImage(
            CreateMinimalRgbPng(),
            24,
            12,
            alternativeText: " "));
    }

    [Fact]
    public void DrawingAlternativeTextRejectsWhitespaceAndDecorativeConflict() {
        Assert.Throws<ArgumentException>(() => new PdfDrawingStyle {
            AlternativeText = " "
        });

        Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().Shape(CreateComplianceShape(), style: new PdfDrawingStyle {
                AlternativeText = "Meaningful badge",
                Decorative = true
            }));
    }
}
