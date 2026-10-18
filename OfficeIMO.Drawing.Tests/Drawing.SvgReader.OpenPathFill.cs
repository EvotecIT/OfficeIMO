using System.Text;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests {
    public sealed class DrawingSvgOpenPathFillTests {
        [Fact]
        public void SvgReaderDefaultsOpenPathFillToBlackAndKeepsItsStrokeOpen() {
            const string path = "<path d='M8 8 H88 V88' stroke='blue' stroke-width='4'/>";

            OfficeRasterImage filled = RenderSvg(path);
            Assert.Equal(OfficeColor.Black, filled.GetPixel(64, 32));
            Assert.Equal(OfficeColor.Blue, filled.GetPixel(87, 48));
            Assert.Equal(OfficeColor.Black, filled.GetPixel(48, 47));

            OfficeRasterImage unfilled = RenderSvg(path.Replace("stroke='blue'", "fill='none' stroke='blue'"));
            Assert.Equal(OfficeColor.White, unfilled.GetPixel(64, 32));
            Assert.Equal(OfficeColor.Blue, unfilled.GetPixel(87, 48));
            Assert.Equal(OfficeColor.White, unfilled.GetPixel(48, 47));
        }

        [Theory]
        [InlineData("nonzero", false)]
        [InlineData("evenodd", true)]
        public void SvgReaderFillsCompoundOpenSubpathsUsingTheAuthoredFillRule(string rule, bool hasHole) {
            string path = "<path d='M8 8 H88 V88 H8 M24 24 H72 V72 H24' fill='red' fill-rule='" + rule + "'/>";

            OfficeRasterImage image = RenderSvg(path);

            Assert.Equal(OfficeColor.Red, image.GetPixel(16, 48));
            Assert.Equal(hasHole ? OfficeColor.White : OfficeColor.Red, image.GetPixel(48, 48));
            Assert.Equal(OfficeColor.White, image.GetPixel(4, 48));
        }

        [Fact]
        public void SvgReaderFillsAnOpenCubicWithTwoEndpointsWithoutStrokingItsClosingChord() {
            OfficeRasterImage image = RenderSvg("<path d='M8 80 C8 8 88 8 88 80' fill='red' stroke='blue' stroke-width='4'/>");

            Assert.Equal(OfficeColor.Red, image.GetPixel(48, 50));
            Assert.Equal(OfficeColor.Red, image.GetPixel(48, 79));
            Assert.Equal(OfficeColor.Blue, image.GetPixel(48, 26));
            Assert.Equal(OfficeColor.White, image.GetPixel(48, 84));
        }

        private static OfficeRasterImage RenderSvg(string content) {
            string source = "<svg xmlns='http://www.w3.org/2000/svg' width='96' height='96'>" + content + "</svg>";
            Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(source), out OfficeDrawing? drawing, out int unsupported));
            Assert.NotNull(drawing);
            Assert.Equal(0, unsupported);
            return OfficeDrawingRasterRenderer.Render(drawing!, 1D, OfficeColor.White);
        }
    }
}
