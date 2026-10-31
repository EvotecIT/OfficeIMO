using System;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfPaginationResourceBoundaryTests {
    [Theory]
    [InlineData(false, 320D)]
    [InlineData(true, 320D)]
    [InlineData(false, 320.0005D)]
    [InlineData(true, 320.0005D)]
    public void ALineAtThePageHeightBoundaryMakesProgressOrIsRejected(bool list, double height) {
        var document = PdfDocument.Create(new PdfCreateOptions {
            PageWidth = OfficeLength.Points(400), PageHeight = OfficeLength.Points(400), MarginTop = OfficeLength.Points(40), MarginBottom = OfficeLength.Points(40),
            MarginLeft = OfficeLength.Points(40), MarginRight = OfficeLength.Points(40), MaxGeneratedPages = 3
        });
        if (list) {
            document.RichNumbered(new[] { new PdfListItem("x") }, style: new PdfListStyle {
                LineSpacing = PdfLineSpacing.Exact(OfficeLength.Points(height)), SpacingBefore = OfficeLength.Points(0), SpacingAfter = OfficeLength.Points(0), ItemSpacing = OfficeLength.Points(0)
            });
        } else {
            document.Paragraph(paragraph => paragraph.Text("x"), style: new PdfParagraphStyle {
                LineSpacing = PdfLineSpacing.Exact(OfficeLength.Points(height)), SpacingBefore = OfficeLength.Points(0), SpacingAfter = OfficeLength.Points(0), WidowControl = false
            });
        }

        if (height > 320D) {
            var error = Assert.Throws<ArgumentException>(() => document.ToBytes());
            Assert.Contains("line height", error.Message, StringComparison.OrdinalIgnoreCase);
        } else {
            Assert.Single(PdfReadDocument.Open(document.ToBytes()).Pages);
        }
    }
}
