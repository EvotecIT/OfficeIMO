using System;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf {
    public partial class PdfDocumentBuilderPageOptionsTests {
        [Fact]
        public void HeaderFooterCompose_RejectsInvalidTypographyAndPlacementValues() {
            var headerFontException = Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page =>
                    page.Header(header => header.Font((PdfStandardFont)99)))));

            Assert.Equal("HeaderFont", headerFontException.ParamName);
            Assert.Contains("PDF header font must be one of the supported standard PDF fonts.", headerFontException.Message, StringComparison.Ordinal);

            var footerFontException = Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page =>
                    page.Footer(footer => footer.Font((PdfStandardFont)99)))));

            Assert.Equal("FooterFont", footerFontException.ParamName);
            Assert.Contains("PDF footer font must be one of the supported standard PDF fonts.", footerFontException.Message, StringComparison.Ordinal);





            var footerSizeException = Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page =>
                    page.Footer(footer => footer.FontSize(global::OfficeIMO.OfficeLength.Points(0))))));

            Assert.Equal("size", footerSizeException.ParamName);





            var footerOffsetException = Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page =>
                    page.Footer(footer => footer.Offset(global::OfficeIMO.OfficeLength.Points(-1))))));

            Assert.Equal("points", footerOffsetException.ParamName);

            var renderHeaderOffsetException = Assert.Throws<ArgumentException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => {
                    page.Margin(global::OfficeIMO.OfficeLength.Points(20));
                    page.Header(header => header.Offset(global::OfficeIMO.OfficeLength.Points(21)).Text("Invalid header offset"));
                    page.Content(content =>
                        content.Column(column =>
                            column.Item().Paragraph(p => p.Text("Body"))));
                })).ToBytes());

            Assert.Contains("PDF header offset must not exceed the top margin when header content is enabled.", renderHeaderOffsetException.Message, StringComparison.Ordinal);

            var renderFooterOffsetException = Assert.Throws<ArgumentException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => {
                    page.Margin(global::OfficeIMO.OfficeLength.Points(20));
                    page.Footer(footer => footer.Offset(global::OfficeIMO.OfficeLength.Points(21)).PageNumber());
                    page.Content(content =>
                        content.Column(column =>
                            column.Item().Paragraph(p => p.Text("Body"))));
                })).ToBytes());

            Assert.Contains("PDF footer offset must not exceed the bottom margin when footer content is enabled.", renderFooterOffsetException.Message, StringComparison.Ordinal);
        }

    }
}
