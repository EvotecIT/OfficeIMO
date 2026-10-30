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
        public void ComposePage_RejectsInvalidPageSetupScalarsAtAssignment() {
            var pageWidthException = Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.Size(0, 792))));
            Assert.Equal("width", pageWidthException.ParamName);

            var pageHeightException = Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.Size(612, double.NaN))));
            Assert.Equal("height", pageHeightException.ParamName);

            var pageSizeException = Assert.Throws<ArgumentOutOfRangeException>(() =>
                new PdfPageSize(OfficeIMO.OfficeLength.Points(612),OfficeIMO.OfficeLength.Points(0)));
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
                new PdfPageMargins(OfficeIMO.OfficeLength.Points(10),OfficeIMO.OfficeLength.Points(-1),OfficeIMO.OfficeLength.Points(10),OfficeIMO.OfficeLength.Points(10)));
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
            AssertMargins(PdfPageMargins.Normal, 72, 72, 72, 72);
            AssertMargins(PdfPageMargins.Narrow, 36, 36, 36, 36);
            AssertMargins(PdfPageMargins.Moderate, 54, 72, 54, 72);
            AssertMargins(PdfPageMargins.Wide, 144, 72, 144, 72);
            AssertMargins(PdfPageMargins.Mirrored, 90, 72, 72, 72);
            AssertMargins(PdfPageMargins.Office2003Default, 90, 72, 90, 72);
            AssertMargins(PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(24)), 24, 24, 24, 24);
        }

        [Fact]
        public void PageSetupValues_CanBeCreatedFromOfficeFriendlyUnits() {
            var letter = new OfficeIMO.Pdf.PdfPageSize(OfficeIMO.OfficeLength.Inches(8.5),OfficeIMO.OfficeLength.Inches(11));
            Assert.Equal(612, letter.Width.ToPoints());
            Assert.Equal(792, letter.Height.ToPoints());

            var a4 = new OfficeIMO.Pdf.PdfPageSize(OfficeIMO.OfficeLength.Centimeters(21),OfficeIMO.OfficeLength.Centimeters(29.7));
            Assert.InRange(a4.Width.ToPoints(), 595.2, 595.4);
            Assert.InRange(a4.Height.ToPoints(), 841.8, 842.0);

            AssertMargins(PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Inches(0.5)), 36, 36, 36, 36);

            var customInches = new OfficeIMO.Pdf.PdfPageMargins(OfficeIMO.OfficeLength.Inches(1),OfficeIMO.OfficeLength.Inches(1.25),OfficeIMO.OfficeLength.Inches(1.5),OfficeIMO.OfficeLength.Inches(2));
            AssertMargins(customInches, 72, 90, 108, 144);

            var customCentimeters = new OfficeIMO.Pdf.PdfPageMargins(OfficeIMO.OfficeLength.Centimeters(2.54),OfficeIMO.OfficeLength.Centimeters(1.27),OfficeIMO.OfficeLength.Centimeters(3.81),OfficeIMO.OfficeLength.Centimeters(5.08));
            AssertMargins(customCentimeters, 72, 36, 108, 144);

            var sizeException = Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeIMO.Pdf.PdfPageSize(OfficeIMO.OfficeLength.Inches(0),OfficeIMO.OfficeLength.Inches(11)));
            Assert.Equal("width", sizeException.ParamName);

        }

        [Fact]
        public void PdfOptionsPageSetupProperties_ApplyReusableValuesAndValidate() {
            var options = new PdfCreateOptions {
                PageSize = PdfPageSizes.A4.Landscape(),
                Margins = PdfPageMargins.Moderate
            };

            Assert.InRange(options.PageWidth.ToPoints(), 841.0, 843.0);
            Assert.InRange(options.PageHeight.ToPoints(), 594.0, 596.0);
            Assert.Equal(OfficePageOrientation.Landscape, options.PageOrientation);
            AssertMargins(options.Margins, 54, 72, 54, 72);

            var exception = Assert.Throws<ArgumentOutOfRangeException>(() =>
                new PdfCreateOptions { PageSize = default });

            Assert.Equal("PageSize", exception.ParamName);
        }

        [Fact]
        public void PdfDocumentPageSetupFluent_AppliesToTopLevelFlowAndComposePages() {
            var doc = PdfDocument.Create()
                .Size(PdfPageSizes.A5)
                .Landscape()
                .Margin(PdfPageMargins.Narrow)
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
                page.Size(PdfPageSizes.A4);
                page.Margin(PdfPageMargins.Narrow);
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
