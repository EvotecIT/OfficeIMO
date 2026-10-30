using System;
using System.Linq;
using OfficeIMO.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests.Pdf {
    public sealed class PdfContainerClosingPaddingTests {
        [Theory]
        [InlineData(45D, 30D, 28D)]
        [InlineData(45D, 30D, 50D)]
        [InlineData(60D, 30D, 28D)]
        public void JoinedSeparatorRestoresOnlyConsumedClosingPadding(double bodyHeight, double padding, double lineAdvance) {
            using var pdf = PdfPigDocument.Open(CreateFixture(bodyHeight, padding, lineAdvance));
            Assert.Equal(2, pdf.NumberOfPages);
            Assert.Equal("First", pdf.GetPage(1).Text);
            Assert.Equal("Next", pdf.GetPage(2).Text);
            var fill = Assert.Single(pdf.GetPage(1).Paths.Where(path => path.IsFilled)
                .Select(path => path.GetBoundingRectangle()), bounds => bounds.HasValue);
            double contentBottom = bodyHeight + 20D - lineAdvance;
            Assert.Equal(contentBottom, fill!.Value.Bottom, 3);
            var sides = pdf.GetPage(1).Paths.Where(path => path.IsStroked)
                .Select(path => path.GetBoundingRectangle())
                .Where(bounds => bounds.HasValue && bounds.Value.Height > .001D).ToArray();
            Assert.Equal(2, sides.Length);
            Assert.All(sides, bounds => Assert.Equal(contentBottom, bounds!.Value.Bottom, 3));
        }

        internal static byte[] CreateFixture(double bodyHeight, double padding, double lineAdvance) {
            var options = new PdfCreateOptions { BackgroundColor = null,
                PageWidth = global::OfficeIMO.OfficeLength.Points(200D), PageHeight = global::OfficeIMO.OfficeLength.Points(bodyHeight + 40D), DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10D),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(20D), MarginRight = global::OfficeIMO.OfficeLength.Points(20D), MarginTop = global::OfficeIMO.OfficeLength.Points(20D), MarginBottom = global::OfficeIMO.OfficeLength.Points(20D)
            };
            var paragraphStyle = new PdfParagraphStyle {
                FontSize = global::OfficeIMO.OfficeLength.Points(10D), SpacingBefore = global::OfficeIMO.OfficeLength.Points(0D), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0D), WidowControl = false,
                LineSpacing = PdfLineSpacing.Multiple(lineAdvance / 10D).WithFontLineBoxBaseline(1.4D)
            };
            var firstStyle = new PdfPanelStyle {
                PaddingX = 0D, PaddingY = 0D, PaddingBottomOverride = padding,
                SpacingBefore = 0D, SpacingAfter = 0D, BorderContinuesWithNext = true,
                Background = global::OfficeIMO.OfficeColor.FromRgb(255, 255, 0), BorderColor = global::OfficeIMO.OfficeColor.Black
            };
            var nextStyle = new PdfPanelStyle {
                PaddingX = 0D, PaddingY = 0D, SpacingBefore = 0D, SpacingAfter = 0D,
                FrameStartPaddingTopOverride = 0D, KeepTogether = true
            };
            return PdfDocument.Create(options)
                .Panel(panel => panel.Paragraph(p => p.Text("First"), style: paragraphStyle), firstStyle)
                .Panel(panel => panel.Paragraph(p => p.Text("Next"), style: new PdfParagraphStyle {
                    FontSize = global::OfficeIMO.OfficeLength.Points(10D), SpacingBefore = global::OfficeIMO.OfficeLength.Points(0D), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0D), LineSpacing = PdfLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(10D))
                }), nextStyle).ToBytes();
        }
    }
}
