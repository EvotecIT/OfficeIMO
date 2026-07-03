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
        public void ComposePage_RejectsInvalidPageSetupScalarsAtAssignment() {
            var pageWidthException = Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.Size(0, 792))));
            Assert.Equal("width", pageWidthException.ParamName);

            var pageHeightException = Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.Size(612, double.NaN))));
            Assert.Equal("height", pageHeightException.ParamName);

            var pageSizeException = Assert.Throws<ArgumentOutOfRangeException>(() =>
                new PageSize(612, double.PositiveInfinity));
            Assert.Equal("height", pageSizeException.ParamName);

            var defaultPageSizeException = Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.Size(default))));
            Assert.Equal("size", defaultPageSizeException.ParamName);

            var uniformMarginException = Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.Margin(-1))));
            Assert.Equal("all", uniformMarginException.ParamName);

            var sideMarginException = Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.Margin(10, 20, double.NegativeInfinity, 20))));
            Assert.Equal("right", sideMarginException.ParamName);

            var pageMarginsException = Assert.Throws<ArgumentOutOfRangeException>(() =>
                new PageMargins(10, double.NaN, 10, 10));
            Assert.Equal("top", pageMarginsException.ParamName);

            var documentPageNumberException = Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfDocument.Create().PageNumberStart(0));
            Assert.Equal("PageNumberStart", documentPageNumberException.ParamName);

            var sectionPageNumberException = Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfDocument.Create().Section(section => section.PageNumberStart(0)));
            Assert.Equal("PageNumberStart", sectionPageNumberException.ParamName);

            var pageNumberStyleException = Assert.Throws<ArgumentException>(() =>
                PdfDocument.Create().PageNumberStyle((PdfPageNumberStyle)99));
            Assert.Equal("PageNumberStyle", pageNumberStyleException.ParamName);
        }

        [Fact]
        public void ComposePage_AllowsMarginsBeforeLargerPageSizeAndKeepsImpossibleFrameRenderTime() {
            var doc = PdfDocument.Create();
            doc.Compose(c => c.Page(page => {
                page.Margin(400);
                page.Size(1000, 1000);
                page.Content(content =>
                    content.Column(column =>
                        column.Item().Paragraph(p => p.Text("Large page after margins."))));
            }));

            using var pdf = PdfPigDocument.Open(new MemoryStream(doc.ToBytes()));
            string text = Normalize(pdf.GetPage(1).Text);
            Assert.Contains("Largepageaftermargins", text, StringComparison.OrdinalIgnoreCase);

            var exception = Assert.Throws<ArgumentException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => {
                    page.Size(200, 200);
                    page.Margin(100);
                    page.Content(content =>
                        content.Column(column =>
                            column.Item().Paragraph(p => p.Text("Impossible frame."))));
                })).ToBytes());

            Assert.Contains("PDF margins must leave a positive content width.", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void PageMarginsPresets_MatchWordCompatiblePointValues() {
            AssertMargins(PageMargins.Normal, 72, 72, 72, 72);
            AssertMargins(PageMargins.Narrow, 36, 36, 36, 36);
            AssertMargins(PageMargins.Moderate, 54, 72, 54, 72);
            AssertMargins(PageMargins.Wide, 144, 72, 144, 72);
            AssertMargins(PageMargins.Mirrored, 90, 72, 72, 72);
            AssertMargins(PageMargins.Office2003Default, 90, 72, 90, 72);
            AssertMargins(PageMargins.Uniform(24), 24, 24, 24, 24);
        }

        [Fact]
        public void PageSetupValues_CanBeCreatedFromOfficeFriendlyUnits() {
            var letter = PageSize.FromInches(8.5, 11);
            Assert.Equal(612, letter.Width);
            Assert.Equal(792, letter.Height);

            var a4 = PageSize.FromCentimeters(21, 29.7);
            Assert.InRange(a4.Width, 595.2, 595.4);
            Assert.InRange(a4.Height, 841.8, 842.0);

            AssertMargins(PageMargins.UniformInches(0.5), 36, 36, 36, 36);

            var customInches = PageMargins.FromInches(1, 1.25, 1.5, 2);
            AssertMargins(customInches, 72, 90, 108, 144);

            var customCentimeters = PageMargins.FromCentimeters(2.54, 1.27, 3.81, 5.08);
            AssertMargins(customCentimeters, 72, 36, 108, 144);

            var sizeException = Assert.Throws<ArgumentOutOfRangeException>(() => PageSize.FromInches(0, 11));
            Assert.Equal("width", sizeException.ParamName);

            var marginException = Assert.Throws<ArgumentOutOfRangeException>(() => PageMargins.UniformCentimeters(double.NaN));
            Assert.Equal("centimeters", marginException.ParamName);
        }

        [Fact]
        public void PdfOptionsPageSetupProperties_ApplyReusableValuesAndValidate() {
            var options = new PdfOptions {
                PageSize = PageSizes.A4.Landscape(),
                Margins = PageMargins.Moderate
            };

            Assert.InRange(options.PageWidth, 841.0, 843.0);
            Assert.InRange(options.PageHeight, 594.0, 596.0);
            Assert.Equal(PdfPageOrientation.Landscape, options.PageOrientation);
            AssertMargins(options.Margins, 54, 72, 54, 72);

            var exception = Assert.Throws<ArgumentOutOfRangeException>(() =>
                new PdfOptions { PageSize = default });

            Assert.Equal("PageSize", exception.ParamName);
        }

        [Fact]
        public void PdfDocumentPageSetupFluent_AppliesToTopLevelFlowAndComposePages() {
            var doc = PdfDocument.Create()
                .Size(PageSizes.A5)
                .Landscape()
                .Margin(PageMargins.Narrow)
                .Paragraph(p => p.Text("Document default page setup."));

            doc.Compose(c => c.Page(page =>
                page.Content(content =>
                    content.Column(column =>
                        column.Item().Paragraph(p => p.Text("Composed page inherits setup."))))));

            using var pdf = PdfPigDocument.Open(new MemoryStream(doc.ToBytes()));
            Assert.Equal(2, pdf.NumberOfPages);

            var page1 = pdf.GetPage(1);
            var page2 = pdf.GetPage(2);

            Assert.InRange(page1.Width, 594.0, 596.0);
            Assert.InRange(page1.Height, 419.0, 421.0);
            Assert.InRange(page2.Width, 594.0, 596.0);
            Assert.InRange(page2.Height, 419.0, 421.0);
            Assert.InRange(FindWordStartX(page1, "Document"), 35.5, 36.5);
            Assert.InRange(FindWordStartX(page2, "Composed"), 35.5, 36.5);
            Assert.Contains("Documentdefaultpagesetup", Normalize(page1.Text), StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Composedpageinheritssetup", Normalize(page2.Text), StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void ComposePageMarginPreset_AppliesReusableMarginValue() {
            var doc = PdfDocument.Create();

            doc.Compose(c => c.Page(page => {
                page.Size(PageSizes.A4);
                page.Margin(PageMargins.Narrow);
                page.Content(content =>
                    content.Column(column =>
                        column.Item().Paragraph(p => p.Text("Narrow margin body."))));
            }));

            using var pdf = PdfPigDocument.Open(new MemoryStream(doc.ToBytes()));
            var firstBodyLetter = pdf.GetPage(1).Letters
                .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
                .OrderByDescending(letter => letter.StartBaseLine.Y)
                .ThenBy(letter => letter.StartBaseLine.X)
                .First();

            Assert.InRange(firstBodyLetter.StartBaseLine.X, 35.5, 36.5);
            Assert.Contains("Narrowmarginbody", Normalize(pdf.GetPage(1).Text), StringComparison.OrdinalIgnoreCase);
        }

    }
}
