using OfficeIMO.Drawing;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests {
    public class DrawingImageComposerTests {
        [Fact]
        public void OfficeImageComposer_ComposesRasterLayersWithAdornments() {
            OfficeRasterImage layer = new OfficeRasterImage(4, 4, OfficeColor.Red);

            byte[] png = OfficeImageComposer.ComposePng(
                12,
                10,
                OfficeColor.White,
                new[] { OfficeImageLayer.FromRaster(layer, 3, 2, 4, 4) },
                beforeLayers: canvas => canvas.FillRectangle(0, 0, 2, 2, OfficeColor.Blue),
                afterLayers: canvas => canvas.FillRectangle(10, 8, 2, 2, OfficeColor.Lime));

            Assert.True(OfficePngReader.TryDecode(png, out OfficeRasterImage? composed));
            Assert.NotNull(composed);
            Assert.Equal(OfficeColor.Blue, composed!.GetPixel(0, 0));
            Assert.Equal(OfficeColor.Red, composed.GetPixel(4, 3));
            Assert.Equal(OfficeColor.Lime, composed.GetPixel(11, 9));
            Assert.Equal(OfficeColor.White, composed.GetPixel(8, 1));
        }

        [Fact]
        public void OfficeImageComposer_ComposesSvgLayersWithRootAndAdornments() {
            string svg = OfficeImageComposer.ComposeSvg(
                20,
                12,
                OfficeColor.FromRgb(240, 240, 240),
                new[] { OfficeImageLayer.FromSvgInner("<circle cx=\"2\" cy=\"2\" r=\"2\"/>", 5, 3, 8, 8) },
                beforeLayers: builder => builder.Append("<text x=\"1\" y=\"2\">Header</text>"),
                afterLayers: builder => builder.Append("<text x=\"1\" y=\"11\">Footer</text>"));

            Assert.StartsWith("<svg xmlns=\"http://www.w3.org/2000/svg\"", svg);
            Assert.Contains("width=\"20\"", svg);
            Assert.Contains("height=\"12\"", svg);
            Assert.Contains("fill=\"#F0F0F0\"", svg);
            Assert.Contains("<text x=\"1\" y=\"2\">Header</text><svg x=\"5\" y=\"3\" width=\"8\" height=\"8\"", svg);
            Assert.Contains("<circle cx=\"2\" cy=\"2\" r=\"2\"/>", svg);
            Assert.EndsWith("<text x=\"1\" y=\"11\">Footer</text></svg>", svg);

            byte[] bytes = OfficeImageComposer.ComposeSvgBytes(20, 12, OfficeColor.White, Array.Empty<OfficeImageLayer>());
            Assert.Equal(Encoding.UTF8.GetString(bytes), OfficeImageComposer.ComposeSvg(20, 12, OfficeColor.White, Array.Empty<OfficeImageLayer>()));
        }
    }
}
