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
            PageWidth = 400, PageHeight = 400, MarginTop = 40, MarginBottom = 40,
            MarginLeft = 40, MarginRight = 40, MaxGeneratedPages = 3
        });
        if (list) {
            document.RichNumbered(new[] { new PdfListItem("x") }, style: new PdfListStyle {
                LineSpacing = PdfLineSpacing.Exactly(height), SpacingBefore = 0, SpacingAfter = 0, ItemSpacing = 0
            });
        } else {
            document.Paragraph(paragraph => paragraph.Text("x"), style: new PdfParagraphStyle {
                LineSpacing = PdfLineSpacing.Exactly(height), SpacingBefore = 0, SpacingAfter = 0, WidowControl = false
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
