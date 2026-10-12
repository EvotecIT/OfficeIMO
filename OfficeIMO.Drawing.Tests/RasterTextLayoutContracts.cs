using System;
using System.IO;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests {
    public sealed class RasterTextLayoutContracts {
        [Theory]
        [InlineData(16D, 1.2D, false, false)]
        [InlineData(16D, 1.2D, false, true)]
        [InlineData(17.5D, 1.13D, false, true)]
        [InlineData(16D, 1.2D, true, true)]
        public void DrawingInMeasuredBoundsRetainsEveryFractionallySpacedLine(double fontSize, double lineHeight, bool wrap, bool clip) {
            OfficeRasterTextOptions options = CreateOptions();
            options.FontSize = fontSize;
            options.LineHeight = lineHeight;
            options.Wrap = wrap;
            options.Clip = clip;
            string text = wrap ? "one two three four" : "A\nB";
            OfficeTextBlockLayout layout = OfficeRasterText.Measure(text, options, wrap ? 64D : null);
            Assert.True(layout.Lines.Count > 1);
            Assert.Equal(fontSize * lineHeight, layout.LineHeight);
            double width = wrap ? 64D : layout.Width;

            OfficeRasterImage expected = RenderMeasuredLayout(layout, options, width, layout.Height);
            OfficeRasterImage actual = new OfficeRasterImage(220, 130, OfficeColor.White);
            OfficeRasterText.Draw(actual, text, 12, 12, width, layout.Height, OfficeColor.Black, options);

            Assert.True(HasInk(expected, 12 + (int)Math.Ceiling(layout.LineHeight), expected.Height));
            Assert.Equal(expected.GetPixels(), actual.GetPixels());
        }

        [Theory]
        [InlineData("ABCDEFGHIJKLMNO\nX")]
        [InlineData("ABCDEFGHIJKLMNO\r\nX")]
        [InlineData("ABCDEFGHIJKLMNO\rX")]
        [InlineData("  A\tB\n\nC\n")]
        public void DisabledWrappingPreservesAuthoredLinesAndWhitespace(string text) {
            OfficeRasterTextOptions options = CreateOptions();
            OfficeTextBlockLayout layout = OfficeRasterText.Measure(text, options);
            OfficeRasterImage expected = RenderMeasuredLayout(layout, options, 64, 100);
            OfficeRasterImage actual = new OfficeRasterImage(220, 130, OfficeColor.White);
            OfficeRasterText.Draw(actual, text, 12, 12, 64, 100, OfficeColor.Black, options);

            Assert.True(HasInk(expected, 12, expected.Height));
            Assert.Equal(expected.GetPixels(), actual.GetPixels());
        }

        [Theory]
        [InlineData(false, false)]
        [InlineData(true, false)]
        [InlineData(false, true)]
        [InlineData(true, true)]
        public void UnclippedDrawingRetainsOverflowingGlyphsAndEffects(bool shadow, bool outline) {
            OfficeRasterTextOptions options = CreateOptions();
            AddEffects(options, shadow, outline);
            OfficeRasterImage ample = new OfficeRasterImage(220, 130, OfficeColor.White);
            OfficeRasterImage shortRectangle = ample.Clone();
            OfficeRasterText.Draw(ample, "A\nB", 12, 12, 100, 80, OfficeColor.Black, options);
            OfficeRasterText.Draw(shortRectangle, "A\nB", 12, 12, 100, 20, OfficeColor.Black, options);

            Assert.True(HasInk(ample, 32, ample.Height));
            Assert.Equal(ample.GetPixels(), shortRectangle.GetPixels());
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void ClippingCutsPixelsWithoutDroppingThePartiallyVisibleNextLine(bool effects) {
            OfficeRasterTextOptions options = CreateOptions();
            options.Clip = true;
            AddEffects(options, effects, effects);
            OfficeRasterImage image = new OfficeRasterImage(220, 130, OfficeColor.White);
            OfficeRasterText.Draw(image, "A\nB", 12, 12, 100, 28, OfficeColor.Black, options);

            Assert.True(HasInk(image, 32, 40), "The second authored line should remain visible up to the clip edge.");
            for (int y = 0; y < image.Height; y++) {
                for (int x = 0; x < image.Width; x++) {
                    if (x < 12 || x >= 112 || y < 12 || y >= 40) {
                        Assert.Equal(OfficeColor.White, image.GetPixel(x, y));
                    }
                }
            }
        }

        [Theory]
        [InlineData(OfficeTextAlignment.Center, OfficeTextVerticalAlignment.Center)]
        [InlineData(OfficeTextAlignment.Right, OfficeTextVerticalAlignment.Bottom)]
        public void AlignmentUsesTheMeasuredFractionalBlock(OfficeTextAlignment horizontal, OfficeTextVerticalAlignment vertical) {
            OfficeRasterTextOptions options = CreateOptions();
            options.HorizontalAlignment = horizontal;
            options.VerticalAlignment = vertical;
            OfficeTextBlockLayout layout = OfficeRasterText.Measure("A\nB", options);
            OfficeRasterImage expected = RenderMeasuredLayout(layout, options, 100, 80);
            OfficeRasterImage actual = new OfficeRasterImage(220, 130, OfficeColor.White);
            OfficeRasterText.Draw(actual, "A\nB", 12, 12, 100, 80, OfficeColor.Black, options);

            Assert.True(HasInk(expected, 12, expected.Height));
            Assert.Equal(expected.GetPixels(), actual.GetPixels());
        }

        private static OfficeRasterTextOptions CreateOptions() => new() {
            FontFamily = "Raster text contract",
            Fonts = new OfficeFontFaceCollection().Add("Raster text contract", File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "Carlito-Regular.ttf")))
        };

        private static OfficeRasterImage RenderMeasuredLayout(OfficeTextBlockLayout layout, OfficeRasterTextOptions options, double width, double height) {
            OfficeRasterImage image = new OfficeRasterImage(220, 130, OfficeColor.White);
            OfficeRasterCanvas canvas = new OfficeRasterCanvas(image, fonts: options.Fonts);
            OfficeTextBlockRenderPlan plan = OfficeTextBlockRenderPlan.CreateFromRectangle(layout, 12, 12, width, height,
                options.HorizontalAlignment, options.VerticalAlignment);
            using (options.Clip ? canvas.PushClipRectangle(12, 12, width, height) : null) {
                OfficeTextBlockRenderer.DrawRasterTextBox(canvas, plan, OfficeColor.Black, fontFamily: options.FontFamily);
            }
            return image;
        }

        private static void AddEffects(OfficeRasterTextOptions options, bool shadow, bool outline) {
            if (shadow) {
                options.ShadowColor = OfficeColor.Red;
                options.ShadowOffsetX = 5;
                options.ShadowOffsetY = 3;
            }
            if (outline) {
                options.OutlineColor = OfficeColor.Blue;
                options.OutlineWidth = 2;
            }
        }

        private static bool HasInk(OfficeRasterImage image, int top, int bottom) {
            for (int y = top; y < bottom; y++) {
                for (int x = 0; x < image.Width; x++) {
                    if (image.GetPixel(x, y) != OfficeColor.White) {
                        return true;
                    }
                }
            }
            return false;
        }
    }
}