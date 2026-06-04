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
        public void ComposePagesHaveIndependentOptions() {
            var doc = PdfDocument.Create();
            doc.Compose(c => {
                c.Page(page => {
                    page.Size(PageSizes.A4);
                    page.Margin(36);
                    page.DefaultTextStyle(style => style.Font(PdfStandardFont.Helvetica).FontSize(14));
                    page.Content(content =>
                        content.Column(col =>
                            col.Item().Paragraph(p => p.Text("First page body."))));
                });
                c.Page(page => {
                    page.Size(PageSizes.Letter);
                    page.Margin(18, 24, 30, 36);
                    page.DefaultTextStyle(style => style.Font(PdfStandardFont.TimesRoman).FontSize(11));
                    page.Footer(f => f.PageNumber());
                    page.Content(content =>
                        content.Column(col =>
                            col.Item().Paragraph(p => p.Text("Second page body."))));
                });
            });

            var bytes = doc.ToBytes();
            Assert.NotEmpty(bytes);

            using (var pdf = PdfPigDocument.Open(new MemoryStream(bytes))) {
                Assert.Equal(2, pdf.NumberOfPages);
                var page1 = pdf.GetPage(1);
                var page2 = pdf.GetPage(2);
                Assert.InRange(page1.Width, 594.0, 596.0); // A4 width
                Assert.InRange(page1.Height, 841.0, 843.0); // A4 height
                Assert.InRange(page2.Width, 611.0, 613.0); // Letter width
                Assert.InRange(page2.Height, 791.0, 793.0); // Letter height

                var page1Text = string.Concat(page1.Letters.Select(l => l.Value));
                var page2Text = string.Concat(page2.Letters.Select(l => l.Value));
                var page1TextNormalized = Normalize(page1Text);
                var page2TextNormalized = Normalize(page2Text);
                Assert.Contains("Firstpagebody", page1TextNormalized, StringComparison.OrdinalIgnoreCase);
                Assert.DoesNotContain("Secondpagebody", page1TextNormalized, StringComparison.OrdinalIgnoreCase);
                Assert.Contains("Secondpagebody", page2TextNormalized, StringComparison.OrdinalIgnoreCase);

                double page1Size = page1.Letters.Where(l => !char.IsWhiteSpace(l.Value[0])).Select(l => l.PointSize).First();
                double page2Size = page2.Letters.Where(l => !char.IsWhiteSpace(l.Value[0])).Select(l => l.PointSize).First();
                Assert.InRange(page1Size, 13.5, 14.5);
                Assert.InRange(page2Size, 10.5, 11.5);
            }
        }

        [Fact]
        public void PdfDocumentPage_CreatesPageScopedFlowWithoutComposeWrapper() {
            var doc = PdfDocument.Create()
                .Size(PageSizes.Letter)
                .Margin(PageMargins.Normal)
                .Header(header => header.Text("Document header {page}/{pages}"))
                .Page(page => {
                    page.Size(PageSizes.A5);
                    page.Margin(PageMargins.Narrow);
                    page.Header(header => header.Text("Small page {page}/{pages}"));
                    page.Content(content =>
                        content.Column(column =>
                            column.Item().Paragraph(p => p.Text("Scoped page body."))));
                })
                .Page(page => {
                    page.Landscape();
                    page.Content(content =>
                        content.Column(column =>
                            column.Item().Paragraph(p => p.Text("Inherited landscape body."))));
                });

            using var pdf = PdfPigDocument.Open(new MemoryStream(doc.ToBytes()));
            Assert.Equal(2, pdf.NumberOfPages);

            var page1 = pdf.GetPage(1);
            var page2 = pdf.GetPage(2);
            string page1Text = Normalize(page1.Text);
            string page2Text = Normalize(page2.Text);

            Assert.InRange(page1.Width, 419.0, 421.0);
            Assert.InRange(page1.Height, 594.0, 596.0);
            Assert.InRange(FindWordStartX(page1, "Scoped"), 35.5, 36.5);
            Assert.Contains("Smallpage1/2", page1Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Scopedpagebody", page1Text, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Documentheader", page1Text, StringComparison.OrdinalIgnoreCase);

            Assert.InRange(page2.Width, 791.0, 793.0);
            Assert.InRange(page2.Height, 611.0, 613.0);
            Assert.InRange(FindWordStartX(page2, "Inherited"), 71.5, 72.5);
            Assert.Contains("Documentheader2/2", page2Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Inheritedlandscapebody", page2Text, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void PdfDocumentSection_CreatesSectionScopedFlowAcrossPhysicalPages() {
            var doc = PdfDocument.Create()
                .Section(section => {
                    section.Size(220, 170);
                    section.Margin(24);
                    section.Header(header => header.Text("Small section {page}/{pages}"));
                    section.Content(content =>
                        content.Column(column => {
                            for (int i = 1; i <= 18; i++) {
                                int item = i;
                                column.Item().Paragraph(p => p.Text("Small section item " + item.ToString("D2", System.Globalization.CultureInfo.InvariantCulture)));
                            }
                        }));
                })
                .Section(section => {
                    section.Size(PageSizes.A5.Landscape());
                    section.Margin(PageMargins.Narrow);
                    section.Header(header => header.Text("Wide section {page}/{pages}"));
                    section.Content(content =>
                        content.Column(column =>
                            column.Item().Paragraph(p => p.Text("Wide section body."))));
                });

            using var pdf = PdfPigDocument.Open(new MemoryStream(doc.ToBytes()));
            Assert.True(pdf.NumberOfPages >= 3, "The first section should flow across multiple physical pages.");

            int widePageNumber = Enumerable.Range(1, pdf.NumberOfPages)
                .First(pageNumber => Normalize(pdf.GetPage(pageNumber).Text).Contains("Widesectionbody", StringComparison.OrdinalIgnoreCase));

            Assert.True(widePageNumber > 1);
            for (int pageNumber = 1; pageNumber < widePageNumber; pageNumber++) {
                var smallPage = pdf.GetPage(pageNumber);
                string text = Normalize(smallPage.Text);
                Assert.InRange(smallPage.Width, 219.0, 221.0);
                Assert.InRange(smallPage.Height, 169.0, 171.0);
                Assert.Contains("Smallsection" + pageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + "/" + pdf.NumberOfPages.ToString(System.Globalization.CultureInfo.InvariantCulture), text, StringComparison.OrdinalIgnoreCase);
                Assert.DoesNotContain("Widesection", text, StringComparison.OrdinalIgnoreCase);
            }

            var widePage = pdf.GetPage(widePageNumber);
            string wideText = Normalize(widePage.Text);
            Assert.InRange(widePage.Width, 594.0, 596.0);
            Assert.InRange(widePage.Height, 419.0, 421.0);
            Assert.Contains("Widesection" + widePageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + "/" + pdf.NumberOfPages.ToString(System.Globalization.CultureInfo.InvariantCulture), wideText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Widesectionbody", wideText, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void PageSizeOrientationHelpers_ReturnExpectedDimensionsAndRejectInvalidOrientation() {
            PageSize portrait = PageSizes.A4.Portrait();
            PageSize landscape = PageSizes.A4.Landscape();

            Assert.InRange(portrait.Width, 594.0, 596.0);
            Assert.InRange(portrait.Height, 841.0, 843.0);
            Assert.InRange(landscape.Width, 841.0, 843.0);
            Assert.InRange(landscape.Height, 594.0, 596.0);
            Assert.Equal(landscape.Width, new PageSize(842, 595).Landscape().Width);
            Assert.Equal(landscape.Height, new PageSize(842, 595).Landscape().Height);

            var exception = Assert.Throws<ArgumentException>(() =>
                PageSizes.A4.WithOrientation((PdfPageOrientation)99));

            Assert.Equal("orientation", exception.ParamName);
            Assert.Contains("PDF page orientation must be Portrait or Landscape.", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void ComposePageOrientation_PreservesPageSizeAndRendersIndependentPageGeometry() {
            var doc = PdfDocument.Create();

            doc.Compose(c => {
                c.Page(page => {
                    page.Size(PageSizes.A4);
                    page.Landscape();
                    page.Content(content =>
                        content.Column(column =>
                            column.Item().Paragraph(p => p.Text("Landscape page body."))));
                });
                c.Page(page => {
                    page.Size(PageSizes.A4.Landscape());
                    page.Portrait();
                    page.Content(content =>
                        content.Column(column =>
                            column.Item().Paragraph(p => p.Text("Portrait page body."))));
                });
            });

            using var pdf = PdfPigDocument.Open(new MemoryStream(doc.ToBytes()));
            Assert.Equal(2, pdf.NumberOfPages);

            var landscapePage = pdf.GetPage(1);
            var portraitPage = pdf.GetPage(2);

            Assert.InRange(landscapePage.Width, 841.0, 843.0);
            Assert.InRange(landscapePage.Height, 594.0, 596.0);
            Assert.InRange(portraitPage.Width, 594.0, 596.0);
            Assert.InRange(portraitPage.Height, 841.0, 843.0);
            Assert.Contains("Landscapepagebody", Normalize(landscapePage.Text), StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Portraitpagebody", Normalize(portraitPage.Text), StringComparison.OrdinalIgnoreCase);

            var exception = Assert.Throws<ArgumentException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.Orientation((PdfPageOrientation)99))));

            Assert.Equal("orientation", exception.ParamName);
        }

    }
}
