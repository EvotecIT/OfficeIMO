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
            var options = new PdfOptions {
                PageWidth = 200D, PageHeight = bodyHeight + 40D, DefaultFontSize = 10D,
                MarginLeft = 20D, MarginRight = 20D, MarginTop = 20D, MarginBottom = 20D
            };
            var paragraphStyle = new PdfParagraphStyle {
                FontSize = 10D, SpacingBefore = 0D, SpacingAfter = 0D, WidowControl = false,
                LineSpacing = PdfLineSpacing.Multiple(lineAdvance / 10D).WithFontLineBoxBaseline(1.4D)
            };
            var firstStyle = new PdfPanelStyle {
                PaddingX = 0D, PaddingY = 0D, PaddingBottomOverride = padding,
                SpacingBefore = 0D, SpacingAfter = 0D, BorderContinuesWithNext = true,
                Background = PdfColor.FromRgb(255, 255, 0), BorderColor = PdfColor.Black
            };
            var nextStyle = new PdfPanelStyle {
                PaddingX = 0D, PaddingY = 0D, SpacingBefore = 0D, SpacingAfter = 0D,
                FrameStartPaddingTopOverride = 0D, KeepTogether = true
            };
            return PdfDocument.Create(options)
                .Panel(panel => panel.Paragraph(p => p.Text("First"), style: paragraphStyle), firstStyle)
                .Panel(panel => panel.Paragraph(p => p.Text("Next"), style: new PdfParagraphStyle {
                    FontSize = 10D, SpacingBefore = 0D, SpacingAfter = 0D, LineSpacing = PdfLineSpacing.Exactly(10D)
                }), nextStyle).ToBytes();
        }
    }
}
