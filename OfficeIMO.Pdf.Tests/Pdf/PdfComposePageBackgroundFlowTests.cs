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
        public void PageBackground_RendersBeforePageContent() {
            byte[] pdfBytes = PdfDocument.Create()
                .Background(PdfColor.FromRgb(240, 248, 255))
                .Paragraph(paragraph => paragraph.Text("DocBackgroundMarker"))
                .ToBytes();

            string content = Encoding.ASCII.GetString(pdfBytes);
            int backgroundFill = content.IndexOf("0.941 0.973 1 rg\n0 0 612 792 re f", StringComparison.Ordinal);
            int markerText = content.IndexOf("<446F634261636B67726F756E644D61726B6572>", StringComparison.Ordinal);

            Assert.True(backgroundFill >= 0, "Expected the document background to emit a full-page PDF fill.");
            Assert.True(markerText > backgroundFill, "Expected the page background to render before text content.");
        }

        [Fact]
        public void PageBackground_CanBeOverriddenPerComposedPage() {
            byte[] pdfBytes = PdfDocument.Create(new PdfOptions {
                    BackgroundColor = PdfColor.White
                })
                .Page(page => page
                    .Size(300, 400)
                    .Background(PdfColor.FromRgb(238, 242, 255))
                    .Content(content => content.Item(item => item.Paragraph(paragraph => paragraph.Text("PageBackgroundMarker")))))
                .ToBytes();

            string content = Encoding.ASCII.GetString(pdfBytes);

            Assert.Contains("0.933 0.949 1 rg\n0 0 300 400 re f", content, StringComparison.Ordinal);
            Assert.DoesNotContain("1 1 1 rg\n0 0 300 400 re f", content, StringComparison.Ordinal);
        }

        [Fact]
        public void ComposeContent_ItemAndSpacerProvideDirectWordLikeFlow() {
            byte[] pdfBytes = PdfDocument.Create(new PdfOptions {
                    DefaultFont = PdfStandardFont.Helvetica,
                    DefaultFontSize = 12
                })
                .Compose(compose => compose.Page(page => page
                    .Margin(72)
                    .Content(content => content
                        .Item(item => item
                            .H1("DirectComposeTitle")
                            .Paragraph(paragraph => paragraph.Text("DirectComposeLead")))
                        .Spacer(24)
                        .Column(column => column
                            .Item(item => item.Paragraph(paragraph => paragraph.Text("ColumnComposeTop")))
                            .Spacer(18)
                            .Item(item => item.Paragraph(paragraph => paragraph.Text("ColumnComposeBottom")))))))
                .ToBytes();

            string text = PdfReadDocument.Open(pdfBytes).ExtractText();
            Assert.Contains("DirectComposeTitle", text, StringComparison.Ordinal);
            Assert.Contains("DirectComposeLead", text, StringComparison.Ordinal);
            Assert.Contains("ColumnComposeTop", text, StringComparison.Ordinal);
            Assert.Contains("ColumnComposeBottom", text, StringComparison.Ordinal);

            using var pdf = PdfPigDocument.Open(new MemoryStream(pdfBytes));
            var page = pdf.GetPage(1);
            double leadY = FindWordStartY(page, "DirectComposeLead");
            double columnTopY = FindWordStartY(page, "ColumnComposeTop");
            double columnBottomY = FindWordStartY(page, "ColumnComposeBottom");

            Assert.True(leadY - columnTopY >= 32, $"Expected direct content spacer to preserve visible rhythm. Lead y: {leadY:0.##}, top y: {columnTopY:0.##}.");
            Assert.True(columnTopY - columnBottomY >= 26, $"Expected column spacer to preserve visible rhythm. Top y: {columnTopY:0.##}, bottom y: {columnBottomY:0.##}.");
        }

        [Fact]
        public void ComposeContent_PageBreaksProvideDirectWordLikeFlow() {
            byte[] pdfBytes = PdfDocument.Create(new PdfOptions {
                    DefaultFont = PdfStandardFont.Helvetica,
                    DefaultFontSize = 12
                })
                .Compose(compose => compose.Page(page => page
                    .Margin(72)
                    .Content(content => content
                        .Item(item => item.Paragraph(paragraph => paragraph.Text("DirectPageOne")))
                        .PageBreak()
                        .Column(column => column
                            .Item(item => item.Paragraph(paragraph => paragraph.Text("ColumnPageTwo")))
                            .PageBreak()
                            .Item(item => item.Paragraph(paragraph => paragraph.Text("ColumnPageThree")))))))
                .ToBytes();

            using var pdf = PdfPigDocument.Open(new MemoryStream(pdfBytes));
            Assert.Equal(3, pdf.NumberOfPages);
            Assert.Contains("DirectPageOne", pdf.GetPage(1).Text, StringComparison.Ordinal);
            Assert.DoesNotContain("ColumnPageTwo", pdf.GetPage(1).Text, StringComparison.Ordinal);
            Assert.Contains("ColumnPageTwo", pdf.GetPage(2).Text, StringComparison.Ordinal);
            Assert.DoesNotContain("ColumnPageThree", pdf.GetPage(2).Text, StringComparison.Ordinal);
            Assert.Contains("ColumnPageThree", pdf.GetPage(3).Text, StringComparison.Ordinal);
        }

        [Fact]
        public void ComposeItem_ElementPageBreakProvidesNestedWordLikeFlow() {
            byte[] pdfBytes = PdfDocument.Create(new PdfOptions {
                    DefaultFont = PdfStandardFont.Helvetica,
                    DefaultFontSize = 12
                })
                .Compose(compose => compose.Page(page => page
                    .Margin(72)
                    .Content(content => content
                        .Column(column => column
                            .Item(item => item
                                .Paragraph(paragraph => paragraph.Text("NestedPageOne"))
                                .Element(element => element
                                    .PageBreak()
                                    .Paragraph(paragraph => paragraph.Text("NestedPageTwo"))))))))
                .ToBytes();

            using var pdf = PdfPigDocument.Open(new MemoryStream(pdfBytes));
            Assert.Equal(2, pdf.NumberOfPages);
            Assert.Contains("NestedPageOne", pdf.GetPage(1).Text, StringComparison.Ordinal);
            Assert.DoesNotContain("NestedPageTwo", pdf.GetPage(1).Text, StringComparison.Ordinal);
            Assert.Contains("NestedPageTwo", pdf.GetPage(2).Text, StringComparison.Ordinal);
        }

    }
}
