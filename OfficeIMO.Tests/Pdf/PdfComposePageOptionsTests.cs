using System;
using System.IO;
using System.Linq;
using OfficeIMO.Pdf;
using UglyToad.PdfPig;
using Xunit;

namespace OfficeIMO.Tests.Pdf {
    public class PdfComposePageOptionsTests {
        [Fact]
        public void ComposePagesHaveIndependentOptions() {
            var doc = PdfDoc.Create();
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

            using (var pdf = PdfDocument.Open(new MemoryStream(bytes))) {
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

        private static string Normalize(string text) {
            return new string(text.Where(c => !char.IsWhiteSpace(c)).ToArray());
        }
    }
}
