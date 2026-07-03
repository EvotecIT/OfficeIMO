using System;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf {
    public partial class PdfComposePageOptionsTests {
        [Fact]
        public void ComposePage_RejectsNullConfigurationDelegates() {
            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.DefaultTextStyle((Action<PdfTextStyleCompose>)null!))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.DefaultTextStyle((PdfTextStyle)null!))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.Theme(null!))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.DefaultParagraphStyle(null!))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.DefaultTableStyle((PdfTableStyle)null!))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.DefaultTableStyle((string)null!))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.DefaultHeadingStyle(1, null!))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.DefaultListStyle(null!))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.DefaultPanelStyle(null!))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.DefaultHorizontalRuleStyle(null!))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.DefaultImageStyle(null!))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.DefaultDrawingStyle(null!))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.DefaultRowStyle(null!))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.Content(null!))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.Content(content => content.Item(null!)))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.Content(content => content.Column(null!)))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.Content(content => content.Column(column => column.Item(null!))))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.Content(content => content.Column(column => column.Item().Element(null!))))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.Content(content => content.Row(null!)))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.Header(null!))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.Footer(null!))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Header(null!));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Footer(null!));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Page(null!));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Section(null!));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Compose(c => c.Section(null!)));
        }

        [Fact]
        public void ComposePage_RejectsInvalidDefaultTextStyleFont() {
            var exception = Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page =>
                    page.DefaultTextStyle(style => style.Font((PdfStandardFont)99)))));

            Assert.Equal("font", exception.ParamName);
            Assert.Contains("PDF default font must be one of the supported standard PDF fonts.", exception.Message, StringComparison.Ordinal);
        }

    }
}
