using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests {
    public class DrawingRasterTests {
        [Fact]
        public void OfficePngWriter_EncodesValidRgbaPng() {
            OfficeRasterImage image = new OfficeRasterImage(4, 3, OfficeColor.White);
            OfficeRasterCanvas canvas = new OfficeRasterCanvas(image);

            canvas.FillRectangle(1, 1, 2, 1, OfficeColor.Red);
            byte[] png = OfficePngWriter.Encode(image);

            OfficeImageInfo info = OfficeImageReader.Identify(png);
            Assert.Equal(OfficeImageFormat.Png, info.Format);
            Assert.Equal(4, info.Width);
            Assert.Equal(3, info.Height);
            Assert.Equal(new byte[] { 0x89, 0x50, 0x4E, 0x47 }, png.Take(4).ToArray());
        }

        [Fact]
        public void OfficeRasterCanvas_MeasuresTextDeterministicallyAcrossRepeatedCalls() {
            OfficeRasterCanvas canvas = new OfficeRasterCanvas(new OfficeRasterImage(80, 24, OfficeColor.Transparent));

            double first = canvas.MeasureText("Repeated text", 12);
            double second = canvas.MeasureText("Repeated text", 12);
            double larger = canvas.MeasureText("Repeated text", 18);

            Assert.True(first > 0D);
            Assert.Equal(first, second);
            Assert.True(larger > first, $"Expected larger font size to measure wider. regular={first}, larger={larger}");
        }

        [Fact]
        public void OfficeRasterCanvas_FillsLinearGradientRectangle() {
            OfficeRasterImage image = new OfficeRasterImage(40, 10, OfficeColor.Transparent);
            OfficeRasterCanvas canvas = new OfficeRasterCanvas(image);

            canvas.FillLinearGradientRectangle(
                0,
                0,
                image.Width,
                image.Height,
                OfficeLinearGradient.Horizontal(OfficeColor.Blue, OfficeColor.Lime));

            OfficeColor left = image.GetPixel(1, 5);
            OfficeColor middle = image.GetPixel(20, 5);
            OfficeColor right = image.GetPixel(38, 5);
            Assert.True(left.B > left.G, $"Expected left gradient edge to be blue-ish, got {left}.");
            Assert.True(right.G > right.B, $"Expected right gradient edge to be green-ish, got {right}.");
            Assert.True(middle.B > 40 && middle.G > 40, $"Expected gradient middle to blend both colors, got {middle}.");
        }

        [Fact]
        public void OfficeRasterCanvas_DrawsSharedHatchPatternRectangle() {
            OfficeRasterImage image = new OfficeRasterImage(32, 24, OfficeColor.Transparent);
            OfficeRasterCanvas canvas = new OfficeRasterCanvas(image);

            canvas.DrawHatchPatternRectangle(2, 2, 24, 16, OfficeColor.Red, 6, 1, OfficeHatchPatternKind.Grid);

            Assert.True(AnyAlpha(image, 7, 2, 8, 18));
            Assert.True(AnyAlpha(image, 2, 7, 26, 8));
            Assert.True(CountPaintedPixels(image) > 50);
            Assert.Equal(0, image.GetPixel(30, 20).A);
        }

        [Fact]
        public void OfficeSparklineRenderer_DrawsRasterLineAndColumnSparklines() {
            OfficeRasterImage lineImage = new OfficeRasterImage(80, 28, OfficeColor.Transparent);
            OfficeSparklineRenderer.DrawRaster(
                new OfficeRasterCanvas(lineImage),
                0,
                0,
                80,
                28,
                new[] { 4D, -2D, 8D },
                OfficeSparklineKind.Line,
                new OfficeSparklineStyle {
                    DisplayAxis = true,
                    AxisColor = OfficeColor.Gray,
                    SeriesColor = OfficeColor.Blue,
                    PointStyles = new[] {
                        new OfficeSparklinePointStyle(OfficeColor.Blue, showMarker: true),
                        new OfficeSparklinePointStyle(OfficeColor.Red, showMarker: true),
                        new OfficeSparklinePointStyle(OfficeColor.Lime, showMarker: true)
                    }
                });

            OfficeRasterImage columnImage = new OfficeRasterImage(80, 28, OfficeColor.Transparent);
            OfficeSparklineRenderer.DrawRaster(
                new OfficeRasterCanvas(columnImage),
                0,
                0,
                80,
                28,
                new[] { 4D, -2D, 8D },
                OfficeSparklineKind.Column,
                new OfficeSparklineStyle {
                    DisplayAxis = true,
                    AxisColor = OfficeColor.Gray,
                    PointStyles = new[] {
                        new OfficeSparklinePointStyle(OfficeColor.Blue),
                        new OfficeSparklinePointStyle(OfficeColor.Red),
                        new OfficeSparklinePointStyle(OfficeColor.Lime)
                    }
                });

            Assert.True(CountPaintedPixels(lineImage) > 30);
            Assert.True(CountPixelsNear(lineImage, OfficeColor.Red) > 0);
            Assert.True(CountPaintedPixels(columnImage) > 100);
            Assert.True(CountPixelsNear(columnImage, OfficeColor.Lime) > 0);
        }

        [Fact]
        public void OfficeDataBarRenderer_DrawsResolvedRasterDataBar() {
            OfficeRasterImage image = new OfficeRasterImage(40, 16, OfficeColor.Transparent);

            OfficeDataBarRenderer.DrawRaster(new OfficeRasterCanvas(image), 2, 3, 30, 10, 0.25D, 0.5D, OfficeColor.Blue, verticalInset: 2D);

            Assert.Equal(0, image.GetPixel(8, 7).A);
            Assert.True(image.GetPixel(10, 7).B > 180);
            Assert.True(image.GetPixel(24, 7).B > 180);
            Assert.Equal(0, image.GetPixel(27, 7).A);
            Assert.Equal(0, image.GetPixel(10, 4).A);
        }

        [Fact]
        public void OfficeRasterCanvas_MeasuresEmptyTextAsZero() {
            OfficeRasterCanvas canvas = new OfficeRasterCanvas(new OfficeRasterImage(16, 16, OfficeColor.Transparent));

            Assert.Equal(0D, canvas.MeasureText(null));
            Assert.Equal(0D, canvas.MeasureText(string.Empty));
        }

        [Fact]
        public void OfficeTextLayoutEngine_WrapsWordsAndHardBreaksWithMeasuredWidths() {
            OfficeRasterCanvas canvas = new OfficeRasterCanvas(new OfficeRasterImage(120, 40, OfficeColor.Transparent));

            IReadOnlyList<OfficeTextLine> lines = OfficeTextLayoutEngine.WrapLines(
                "Alpha\tBeta\nGamma",
                12D,
                canvas.MeasureText("Alpha Beta", 12D) - 1D,
                canvas.MeasureText);

            Assert.Equal(3, lines.Count);
            Assert.Equal("Alpha", lines[0].Text);
            Assert.Equal("Beta", lines[1].Text);
            Assert.Equal("Gamma", lines[2].Text);
            Assert.All(lines, line => Assert.True(line.Width <= canvas.MeasureText("Alpha Beta", 12D)));
            Assert.Equal(lines.Max(line => line.Width), OfficeTextLayoutEngine.MeasureMaxLineWidth(lines));
        }

        [Fact]
        public void OfficeTextLayoutEngine_BreaksLongWordsAndReturnsEmptyLineForBlankText() {
            double Measure(string? value, double size) => (value?.Length ?? 0) * size;

            IReadOnlyList<OfficeTextLine> broken = OfficeTextLayoutEngine.WrapLines("ABCDE", 1D, 2D, Measure);
            IReadOnlyList<OfficeTextLine> blank = OfficeTextLayoutEngine.WrapLines(" \r\n ", 1D, 2D, Measure);

            Assert.Equal(new[] { "AB", "CD", "E" }, broken.Select(line => line.Text).ToArray());
            Assert.Equal(2D, OfficeTextLayoutEngine.MeasureMaxLineWidth(broken));
            Assert.Equal(2, blank.Count);
            Assert.All(blank, line => {
                Assert.Equal(string.Empty, line.Text);
                Assert.Equal(0D, line.Width);
            });
        }

        [Fact]
        public void OfficeTextLayoutEngine_TrimsSingleLineWithEllipsisWhenNeeded() {
            double Measure(string? value, double size) => (value?.Length ?? 0) * size;

            OfficeTextLine clipped = OfficeTextLayoutEngine.TrimLineToWidth("ABCDEFG", 1D, 6D, Measure, out bool wasClipped);
            OfficeTextLine unchanged = OfficeTextLayoutEngine.TrimLineToWidth("ABC", 1D, 6D, Measure, out bool unchangedClipped);
            OfficeTextLine startClipped = OfficeTextLayoutEngine.TrimLineStartToWidth("ABCDEFG", 1D, 5D, Measure, out bool wasStartClipped);

            Assert.True(wasClipped);
            Assert.Equal("ABC...", clipped.Text);
            Assert.Equal(6D, clipped.Width);
            Assert.False(unchangedClipped);
            Assert.Equal("ABC", unchanged.Text);
            Assert.Equal(3D, unchanged.Width);
            Assert.True(wasStartClipped);
            Assert.Equal("...FG", startClipped.Text);
            Assert.Equal(5D, startClipped.Width);
        }

        [Fact]
        public void OfficeTextZoneLayout_CreatesNonOverlappingThreeColumnZones() {
            OfficeTextZoneLayout zones = OfficeTextZoneLayout.CreateThreeColumn(144D, 12D, 6D);

            Assert.Equal(12D, zones.Left.X);
            Assert.Equal(36D, zones.Left.Width);
            Assert.Equal(12D, zones.Left.AnchorX);
            Assert.Equal(54D, zones.Center.X);
            Assert.Equal(72D, zones.Center.AnchorX);
            Assert.Equal(96D, zones.Right.X);
            Assert.Equal(132D, zones.Right.AnchorX);
            Assert.True(zones.Left.X + zones.Left.Width < zones.Center.X);
            Assert.True(zones.Center.X + zones.Center.Width < zones.Right.X);
        }

        [Fact]
        public void OfficeTextLayoutEngine_FitsWrappedTextBlockByScalingFontSize() {
            double Measure(string? value, double size) => (value?.Length ?? 0) * size;

            OfficeTextBlockLayout layout = OfficeTextLayoutEngine.FitWrappedText(
                "Alpha\nBeta\nGamma",
                10D,
                100D,
                20D,
                lineHeightFactor: 1.2D,
                minimumFontSize: 4D,
                Measure);

            Assert.Equal(new[] { "Alpha", "Beta", "Gamma" }, layout.Lines.Select(line => line.Text).ToArray());
            Assert.True(layout.FontSize < 10D);
            Assert.True(layout.FontSize > 4D);
            Assert.True(layout.Height <= 20.01D);
            Assert.Equal(layout.Lines.Max(line => line.Width), layout.Width);
        }

        [Fact]
        public void OfficeTextLayoutEngine_ClipsTextBlockToVisibleHeightWithEllipsis() {
            double Measure(string? value, double size) => (value?.Length ?? 0) * size;
            IReadOnlyList<OfficeTextLine> lines = new[] {
                new OfficeTextLine("Alpha", 5D),
                new OfficeTextLine("Beta", 4D),
                new OfficeTextLine("Gamma", 5D)
            };

            OfficeTextBlockLayout layout = OfficeTextLayoutEngine.ClipTextBlockToHeight(
                lines,
                fontSize: 1D,
                lineHeight: 10D,
                maxWidth: 8D,
                maxHeight: 20D,
                Measure);

            Assert.True(layout.Clipped);
            Assert.Equal(new[] { "Alpha", "Beta..." }, layout.Lines.Select(line => line.Text).ToArray());
            Assert.Equal(7D, layout.Width);
            Assert.Equal(20D, layout.Height);
            Assert.Equal(10D, layout.LineHeight);
        }

        [Fact]
        public void OfficeTextLayoutEngine_FitsSingleLineFontSizeWithinBounds() {
            double Measure(string? value, double size) => (value?.Length ?? 0) * size;

            double unchanged = OfficeTextLayoutEngine.FitSingleLineFontSize("ABC", 10D, 30D, 2D, Measure);
            double fitted = OfficeTextLayoutEngine.FitSingleLineFontSize("ABCDE", 10D, 25D, 2D, Measure);
            double floored = OfficeTextLayoutEngine.FitSingleLineFontSize("ABCDE", 10D, 5D, 2D, Measure);

            Assert.Equal(10D, unchanged);
            Assert.InRange(fitted, 4.99D, 5.01D);
            Assert.Equal(2D, floored);
        }

        [Fact]
        public void OfficeTextLayoutEngine_LayoutsBoundedTextBlockWithShrinkWrapAndClipping() {
            double Measure(string? value, double size) => (value?.Length ?? 0) * size;

            OfficeTextBlockLayout shrink = OfficeTextLayoutEngine.LayoutTextBlock(
                "ABCDE",
                10D,
                25D,
                20D,
                lineHeightFactor: 1.2D,
                minimumFontSize: 2D,
                Measure,
                wrap: false,
                shrinkToFit: true);
            OfficeTextBlockLayout wrapped = OfficeTextLayoutEngine.LayoutTextBlock(
                "Alpha Beta Gamma",
                10D,
                50D,
                24D,
                lineHeightFactor: 1.2D,
                minimumFontSize: 2D,
                Measure,
                wrap: true);
            OfficeTextBlockLayout single = OfficeTextLayoutEngine.LayoutTextBlock(
                "A\r\nB",
                10D,
                100D,
                20D,
                lineHeightFactor: 1.2D,
                minimumFontSize: 2D,
                Measure,
                wrap: true,
                forceSingleLine: true);

            Assert.InRange(shrink.FontSize, 4.99D, 5.01D);
            Assert.False(shrink.Clipped);
            Assert.Equal(new[] { "ABCDE" }, shrink.Lines.Select(line => line.Text).ToArray());
            Assert.True(wrapped.Clipped);
            Assert.Equal(2, wrapped.Lines.Count);
            Assert.EndsWith("...", wrapped.Lines[1].Text);
            Assert.Equal(new[] { "A B" }, single.Lines.Select(line => line.Text).ToArray());
        }

        [Fact]
        public void OfficeTextLayoutEngine_LayoutsRichTextRunsWithWrappingHardBreaksAndClipping() {
            double Measure(string? value, double size) => (value?.Length ?? 0) * size;
            OfficeRichTextRun strong = new OfficeRichTextRun("Alpha ", 10D, OfficeColor.Red, bold: true);
            OfficeRichTextRun accent = new OfficeRichTextRun("Beta Gamma", 10D, OfficeColor.Blue, italic: true);
            OfficeRichTextRun hardBreak = new OfficeRichTextRun("A\r\nB", 10D, OfficeColor.Black);

            OfficeRichTextBlockLayout wrapped = OfficeTextLayoutEngine.LayoutRichTextBlock(
                new[] { strong, accent },
                50D,
                24D,
                lineHeightFactor: 1.2D,
                Measure,
                wrap: true);
            OfficeRichTextBlockLayout multiline = OfficeTextLayoutEngine.LayoutRichTextBlock(
                new[] { hardBreak },
                100D,
                30D,
                lineHeightFactor: 1.2D,
                Measure,
                wrap: false);
            OfficeRichTextBlockLayout shrink = OfficeTextLayoutEngine.LayoutRichTextBlock(
                new[] {
                    new OfficeRichTextRun("Wide", 10D, OfficeColor.Red, bold: true),
                    new OfficeRichTextRun(" Text", 20D, OfficeColor.Blue, italic: true)
                },
                70D,
                30D,
                lineHeightFactor: 1.2D,
                Measure,
                wrap: false,
                shrinkToFit: true,
                minimumFontSize: 2D);

            Assert.True(wrapped.Clipped);
            Assert.Equal(2, wrapped.Lines.Count);
            Assert.Equal("Alpha", string.Concat(wrapped.Lines[0].Segments.Select(segment => segment.Text)));
            Assert.True(wrapped.Lines[0].Segments[0].Bold);
            Assert.Equal(OfficeColor.Red, wrapped.Lines[0].Segments[0].Color);
            Assert.EndsWith("...", string.Concat(wrapped.Lines[1].Segments.Select(segment => segment.Text)));
            Assert.True(wrapped.Lines[1].Segments.Last().Italic);
            Assert.Equal(OfficeColor.Blue, wrapped.Lines[1].Segments.Last().Color);
            Assert.Equal(new[] { "A", "B" }, multiline.Lines.Select(line => string.Concat(line.Segments.Select(segment => segment.Text))).ToArray());
            Assert.False(shrink.Clipped);
            Assert.True(shrink.Width <= 70.01D);
            Assert.Equal("Wide Text", string.Concat(shrink.Lines[0].Segments.Select(segment => segment.Text)));
            Assert.True(shrink.Lines[0].Segments[0].FontSize < 10D);
            Assert.True(shrink.Lines[0].Segments[1].FontSize < 20D);
            Assert.True(shrink.Lines[0].Segments[0].Bold);
            Assert.True(shrink.Lines[0].Segments[1].Italic);
        }

        [Fact]
        public void OfficeTextLayoutEngine_LayoutsRichTextRunsWithFontFamilyAwareMeasurement() {
            double Measure(string? value, double size, string? family) {
                double factor = string.Equals(family, "Wide", StringComparison.Ordinal) ? 10D : 1D;
                return (value?.Length ?? 0) * size * factor;
            }

            OfficeRichTextBlockLayout layout = OfficeTextLayoutEngine.LayoutRichTextBlock(
                new[] {
                    new OfficeRichTextRun("AA", 2D, OfficeColor.Black, fontFamily: "Wide"),
                    new OfficeRichTextRun("BB", 2D, OfficeColor.Black, fontFamily: "Narrow")
                },
                100D,
                30D,
                lineHeightFactor: 1.2D,
                Measure,
                wrap: false);

            Assert.Single(layout.Lines);
            Assert.Equal(40D, layout.Lines[0].Segments[0].Width);
            Assert.Equal(4D, layout.Lines[0].Segments[1].Width);
            Assert.Equal(44D, layout.Width);
        }

        [Fact]
        public void OfficeTextPlacement_ResolvesSharedHorizontalAndVerticalCoordinates() {
            Assert.Equal(10D, OfficeTextPlacement.ResolveAnchorX(10D, 100D, OfficeTextAlignment.Left));
            Assert.Equal(60D, OfficeTextPlacement.ResolveAnchorX(10D, 100D, OfficeTextAlignment.Center));
            Assert.Equal(110D, OfficeTextPlacement.ResolveAnchorX(10D, 100D, OfficeTextAlignment.Right));

            Assert.Equal(20D, OfficeTextPlacement.ResolveAnchorXFromCenter(70D, 100D, OfficeTextAlignment.Left));
            Assert.Equal(70D, OfficeTextPlacement.ResolveAnchorXFromCenter(70D, 100D, OfficeTextAlignment.Center));
            Assert.Equal(120D, OfficeTextPlacement.ResolveAnchorXFromCenter(70D, 100D, OfficeTextAlignment.Right));
            Assert.Equal(70D, OfficeTextPlacement.ResolveAnchorXFromCenter(70D, double.PositiveInfinity, OfficeTextAlignment.Right));

            Assert.Equal(80D, OfficeTextPlacement.ResolveLeftFromAnchor(120D, 40D, OfficeTextAlignment.Right));
            Assert.Equal(100D, OfficeTextPlacement.ResolveLeftFromAnchor(120D, 40D, OfficeTextAlignment.Center));
            Assert.Equal(120D, OfficeTextPlacement.ResolveLeftFromAnchor(120D, 40D, OfficeTextAlignment.Left));

            Assert.Equal(10D, OfficeTextPlacement.ResolveLineLeft(10D, 100D, 40D, OfficeTextAlignment.Left));
            Assert.Equal(40D, OfficeTextPlacement.ResolveLineLeft(10D, 100D, 40D, OfficeTextAlignment.Center));
            Assert.Equal(70D, OfficeTextPlacement.ResolveLineLeft(10D, 100D, 40D, OfficeTextAlignment.Right));

            Assert.Equal(10D, OfficeTextPlacement.ResolveTop(10D, 80D, 20D, OfficeTextVerticalAlignment.Top));
            Assert.Equal(40D, OfficeTextPlacement.ResolveTop(10D, 80D, 20D, OfficeTextVerticalAlignment.Center));
            Assert.Equal(70D, OfficeTextPlacement.ResolveTop(10D, 80D, 20D, OfficeTextVerticalAlignment.Bottom));

            Assert.Equal(60D, OfficeTextPlacement.ResolveTopFromCenter(100D, 80D, 20D, OfficeTextVerticalAlignment.Top));
            Assert.Equal(90D, OfficeTextPlacement.ResolveTopFromCenter(100D, 80D, 20D, OfficeTextVerticalAlignment.Center));
            Assert.Equal(120D, OfficeTextPlacement.ResolveTopFromCenter(100D, 80D, 20D, OfficeTextVerticalAlignment.Bottom));
            Assert.Equal(90D, OfficeTextPlacement.ResolveTopFromCenter(100D, double.PositiveInfinity, 20D, OfficeTextVerticalAlignment.Top));

            OfficePoint rotated = OfficeTextPlacement.RotatePoint(new OfficePoint(10D, 0D), 0D, 0D, 90D);
            Assert.InRange(rotated.X, -0.0001D, 0.0001D);
            Assert.InRange(rotated.Y, 9.9999D, 10.0001D);
            Assert.Equal(new OfficePoint(5D, 7D), OfficeTextPlacement.RotatePoint(new OfficePoint(5D, 7D), 0D, 0D, 0D));
        }

        [Fact]
        public void OfficeTextBlockRenderPlan_ResolvesCenteredPlacementAndBackgroundBounds() {
            var layout = new OfficeTextBlockLayout(
                new[] { new OfficeTextLine("Shared", 30D) },
                fontSize: 10D,
                lineHeight: 12D,
                width: 30D,
                height: 12D);

            OfficeTextBlockRenderPlan plan = OfficeTextBlockRenderPlan.CreateFromCenter(
                layout,
                centerX: 100D,
                centerY: 50D,
                width: 80D,
                height: 40D,
                OfficeTextAlignment.Right,
                OfficeTextVerticalAlignment.Bottom);

            Assert.Equal(60D, plan.Left);
            Assert.Equal(30D, plan.Top);
            Assert.Equal(140D, plan.AnchorX);
            Assert.Equal(110D, plan.TextLeft);
            Assert.Equal(58D, plan.TextTop);

            OfficeTextBlockBackgroundBounds background = plan.CreateBackgroundBounds(4D, 2D);
            Assert.Equal(106D, background.Left);
            Assert.Equal(56D, background.Top);
            Assert.Equal(38D, background.Width);
            Assert.Equal(16D, background.Height);

            OfficePoint[] corners = background.GetRotatedCorners(90D, 100D, 50D);
            Assert.Equal(4, corners.Length);
            Assert.InRange(corners[0].X, 93.999D, 94.001D);
            Assert.InRange(corners[0].Y, 55.999D, 56.001D);
        }

        [Fact]
        public void OfficeTextBlockRenderPlan_FitsTextBeforeResolvingPlacement() {
            static double Measure(string? value, double size) => (value?.Length ?? 0) * size * 0.5D;

            OfficeTextBlockRenderPlan plan = OfficeTextBlockRenderPlan.CreateFittedFromCenter(
                "alpha beta gamma",
                fontSize: 12D,
                centerX: 100D,
                centerY: 80D,
                width: 48D,
                height: 30D,
                Measure,
                OfficeTextAlignment.Center,
                OfficeTextVerticalAlignment.Center,
                lineHeightFactor: 1.2D,
                minimumFontSize: 6D);

            Assert.Equal(76D, plan.Left);
            Assert.Equal(65D, plan.Top);
            Assert.Equal(100D, plan.AnchorX);
            Assert.True(plan.Layout.Width <= 48D);
            Assert.True(plan.Layout.Height <= 30D);
        }

        [Fact]
        public void OfficeTextBlockRenderPlan_CreatesRectangleBasedTextAndStackedLayouts() {
            static double Measure(string? value, double size) => (value?.Length ?? 0) * size * 0.8D;

            OfficeTextBlockRenderPlan wrapped = OfficeTextBlockRenderPlan.CreateTextBlockFromRectangle(
                "alpha beta",
                fontSize: 10D,
                left: 20D,
                top: 30D,
                width: 44D,
                height: 50D,
                Measure,
                OfficeTextAlignment.Right,
                OfficeTextVerticalAlignment.Bottom,
                lineHeightFactor: 1.2D,
                minimumFontSize: 6D,
                wrap: true);

            Assert.Equal(20D, wrapped.Left);
            Assert.Equal(30D, wrapped.Top);
            Assert.Equal(64D, wrapped.AnchorX);
            Assert.Equal(56D, wrapped.TextTop);
            Assert.Equal(2, wrapped.Layout.Lines.Count);
            Assert.Equal("alpha", wrapped.Layout.Lines[0].Text);
            Assert.Equal("beta", wrapped.Layout.Lines[1].Text);

            OfficeTextBlockRenderPlan stacked = OfficeTextBlockRenderPlan.CreateStackedTextBlockFromRectangle(
                "AB",
                fontSize: 10D,
                left: 5D,
                top: 7D,
                width: 20D,
                height: 30D,
                Measure,
                OfficeTextAlignment.Center,
                OfficeTextVerticalAlignment.Center);

            Assert.Equal(15D, stacked.AnchorX);
            Assert.Equal(2, stacked.Layout.Lines.Count);
            Assert.Equal("A", stacked.Layout.Lines[0].Text);
            Assert.Equal("B", stacked.Layout.Lines[1].Text);
        }

        [Fact]
        public void OfficeRasterCanvas_ClipsDrawingToRectangleAndRestoresPreviousClip() {
            OfficeRasterImage image = new OfficeRasterImage(12, 12, OfficeColor.Transparent);
            OfficeRasterCanvas canvas = new OfficeRasterCanvas(image);

            using (canvas.PushClipRectangle(3, 3, 6, 6)) {
                canvas.FillRectangle(0, 0, 12, 12, OfficeColor.Red);
                using (canvas.PushClipRectangle(5, 5, 2, 2)) {
                    canvas.FillRectangle(0, 0, 12, 12, OfficeColor.Blue);
                }

                canvas.FillRectangle(8, 8, 1, 1, OfficeColor.Black);
            }

            canvas.FillRectangle(0, 0, 1, 1, OfficeColor.Black);

            Assert.Equal(0, image.GetPixel(2, 2).A);
            Assert.Equal(255, image.GetPixel(3, 3).R);
            Assert.Equal(255, image.GetPixel(5, 5).B);
            Assert.Equal(0, image.GetPixel(8, 8).R);
            Assert.Equal(255, image.GetPixel(0, 0).A);
        }

        [Fact]
        public void OfficeRasterCanvas_DrawsAntialiasedLines() {
            OfficeRasterImage image = new OfficeRasterImage(24, 24, OfficeColor.Transparent);
            OfficeRasterCanvas canvas = new OfficeRasterCanvas(image);

            canvas.DrawLine(2.2, 3.35, 21.4, 18.1, OfficeColor.Black, 1);

            Assert.Contains(Enumerable.Range(0, image.Width * image.Height), index => {
                OfficeColor pixel = image.GetPixel(index % image.Width, index / image.Width);
                return pixel.A > 0 && pixel.A < 255;
            });
        }

        [Fact]
        public void OfficeRasterRenderTarget_ResolvesSupersampledAlphaWeightedPixels() {
            OfficeRasterRenderTarget target = new OfficeRasterRenderTarget(1, 1, supersampling: 2);

            target.SetPixel(0, 0, OfficeColor.FromRgba(255, 0, 0, 255));
            target.SetPixel(1, 0, OfficeColor.FromRgba(0, 0, 255, 255));
            target.SetPixel(0, 1, OfficeColor.Transparent);
            target.SetPixel(1, 1, OfficeColor.Transparent);

            byte[] rgba = target.ResolveRgba();

            Assert.Equal(128, rgba[0]);
            Assert.Equal(0, rgba[1]);
            Assert.Equal(128, rgba[2]);
            Assert.Equal(127, rgba[3]);
        }

        [Fact]
        public void OfficeRasterCanvas_UsesSupersamplingForAntialiasedRenderTargetEdges() {
            OfficeRasterRenderTarget target = new OfficeRasterRenderTarget(12, 12, supersampling: 3);
            OfficeRasterCanvas canvas = new OfficeRasterCanvas(target);

            canvas.FillEllipse(2.2, 2.2, 7.6, 7.6, OfficeColor.Black);
            byte[] rgba = target.ResolveRgba();

            Assert.Contains(Enumerable.Range(0, 12 * 12), index => {
                byte alpha = rgba[(index * 4) + 3];
                return alpha > 0 && alpha < 255;
            });
        }

        [Fact]
        public void OfficeRasterRenderTarget_BlendsPixelsBeforeResolve() {
            OfficeRasterRenderTarget target = new OfficeRasterRenderTarget(1, 1);

            target.SetPixel(0, 0, OfficeColor.FromRgba(0, 0, 255, 255));
            target.BlendPixel(0, 0, OfficeColor.FromRgba(255, 0, 0, 128));
            byte[] rgba = target.ResolveRgba();

            Assert.True(rgba[0] > 120);
            Assert.Equal(0, rgba[1]);
            Assert.True(rgba[2] > 120);
            Assert.Equal(255, rgba[3]);
        }

        [Fact]
        public void OfficeRasterCanvas_CanDrawOnSupersampledRenderTarget() {
            OfficeRasterRenderTarget target = new OfficeRasterRenderTarget(8, 8, supersampling: 2, OfficeColor.Transparent);
            OfficeRasterCanvas canvas = new OfficeRasterCanvas(target);

            canvas.FillPolygon(new[] {
                new OfficePoint(2, 2),
                new OfficePoint(13, 2),
                new OfficePoint(13, 13),
                new OfficePoint(2, 13)
            }, OfficeColor.Black);

            byte[] rgba = target.ResolveRgba();

            Assert.Equal(8 * 8 * 4, rgba.Length);
            Assert.True(rgba[((3 * 8) + 3) * 4 + 3] > 0);
            Assert.Equal(0, rgba[((0 * 8) + 0) * 4 + 3]);
        }

        [Fact]
        public void OfficeRasterCanvas_DrawsDashedLinesWithTransparentGaps() {
            OfficeRasterImage image = new OfficeRasterImage(28, 11, OfficeColor.Transparent);
            OfficeRasterCanvas canvas = new OfficeRasterCanvas(image);

            canvas.DrawDashedLine(2, 5, 26, 5, OfficeColor.Black, 1, dashLength: 4, gapLength: 4);

            Assert.True(image.GetPixel(3, 5).A > 0);
            Assert.Equal(0, image.GetPixel(8, 5).A);
            Assert.True(image.GetPixel(12, 5).A > 0);
        }

        [Fact]
        public void OfficeRasterCanvas_DrawsStyledDashDotDotLines() {
            OfficeRasterImage image = new OfficeRasterImage(64, 14, OfficeColor.Transparent);
            OfficeRasterCanvas canvas = new OfficeRasterCanvas(image);

            canvas.DrawStyledLine(2, 7, 62, 7, OfficeColor.Black, 2, OfficeStrokeDashStyle.DashDotDot);

            Assert.True(AnyAlpha(image, 2, 6, 8, 8));
            Assert.True(CountTransparentColumnsOnRow(image, 7, 2, 62) >= 4);
        }

        [Fact]
        public void OfficeRasterCanvas_DrawsParallelStyledLinesThroughSharedGeometry() {
            OfficeRasterImage image = new OfficeRasterImage(24, 18, OfficeColor.Transparent);
            OfficeRasterCanvas canvas = new OfficeRasterCanvas(image);

            canvas.DrawParallelStyledLine(12, 3, 12, 15, OfficeColor.Black, 1, 6);

            Assert.True(image.GetPixel(9, 9).A > 0);
            Assert.True(image.GetPixel(15, 9).A > 0);
            Assert.Equal(0, image.GetPixel(12, 9).A);
        }

        [Fact]
        public void OfficeDrawingRasterRenderer_HonorsLineDashStyle() {
            OfficeDrawing solidDrawing = new OfficeDrawing(72, 16);
            OfficeShape solidLine = OfficeShape.Line(0, 0, 64, 0);
            solidLine.StrokeColor = OfficeColor.Black;
            solidLine.StrokeWidth = 2;
            solidDrawing.AddShape(solidLine, 4, 8);

            OfficeDrawing dashedDrawing = new OfficeDrawing(72, 16);
            OfficeShape dashedLine = OfficeShape.Line(0, 0, 64, 0);
            dashedLine.StrokeColor = OfficeColor.Black;
            dashedLine.StrokeWidth = 2;
            dashedLine.StrokeDashStyle = OfficeStrokeDashStyle.Dash;
            dashedDrawing.AddShape(dashedLine, 4, 8);

            OfficeRasterImage solid = OfficeDrawingRasterRenderer.Render(solidDrawing);
            OfficeRasterImage dashed = OfficeDrawingRasterRenderer.Render(dashedDrawing);

            Assert.True(CountPaintedPixels(solid) > CountPaintedPixels(dashed));
            Assert.True(AnyAlpha(dashed, 4, 7, 10, 9));
            Assert.True(CountTransparentColumnsOnRow(dashed, 8, 4, 68) >= 6);
        }

        [Fact]
        public void OfficeDrawingRasterRenderer_HonorsPolygonDashStyleThroughSharedRasterPrimitive() {
            OfficeDrawing solidDrawing = new OfficeDrawing(72, 32);
            OfficeShape solidPolygon = OfficeShape.Polygon(
                new OfficePoint(0, 0),
                new OfficePoint(64, 0),
                new OfficePoint(64, 20),
                new OfficePoint(0, 20));
            solidPolygon.StrokeColor = OfficeColor.Black;
            solidPolygon.StrokeWidth = 2;
            solidDrawing.AddShape(solidPolygon, 4, 6);

            OfficeDrawing dashedDrawing = new OfficeDrawing(72, 32);
            OfficeShape dashedPolygon = solidPolygon.Clone();
            dashedPolygon.StrokeDashStyle = OfficeStrokeDashStyle.Dash;
            dashedDrawing.AddShape(dashedPolygon, 4, 6);

            OfficeRasterImage solid = OfficeDrawingRasterRenderer.Render(solidDrawing);
            OfficeRasterImage dashed = OfficeDrawingRasterRenderer.Render(dashedDrawing);

            Assert.True(CountPaintedPixels(solid) > CountPaintedPixels(dashed));
            Assert.True(AnyAlpha(dashed, 4, 5, 12, 7));
            Assert.True(AnyAlpha(dashed, 63, 6, 69, 16));
            Assert.True(CountTransparentColumnsOnRow(dashed, 6, 4, 68) >= 6);
        }

        [Fact]
        public void OfficeDrawingRasterRenderer_HonorsShapeTransformsInRasterOutput() {
            OfficeDrawing untransformedDrawing = new OfficeDrawing(40, 40);
            OfficeShape untransformedShape = OfficeShape.Rectangle(20, 6);
            untransformedShape.FillColor = OfficeColor.Red;
            untransformedDrawing.AddShape(untransformedShape, 10, 17);

            OfficeDrawing transformedDrawing = new OfficeDrawing(40, 40);
            OfficeShape transformedShape = OfficeShape.Rectangle(20, 6);
            transformedShape.FillColor = OfficeColor.Red;
            transformedShape.Transform = OfficeTransform.RotateDegrees(90, 10, 3);
            transformedDrawing.AddShape(transformedShape, 10, 17);

            OfficeRasterImage untransformed = OfficeDrawingRasterRenderer.Render(untransformedDrawing);
            OfficeRasterImage transformed = OfficeDrawingRasterRenderer.Render(transformedDrawing);

            Assert.True(CountPixelsNear(transformed, OfficeColor.Red) > 80);
            Assert.Equal(0, untransformed.GetPixel(20, 11).A);
            AssertColorNear(transformed.GetPixel(20, 11), OfficeColor.Red);
            AssertColorNear(untransformed.GetPixel(12, 20), OfficeColor.Red);
            Assert.Equal(0, transformed.GetPixel(12, 20).A);
        }

        [Fact]
        public void OfficeDrawingRasterRenderer_RendersRotatedTextThroughSharedTextRenderer() {
            OfficeDrawing drawing = new OfficeDrawing(96, 64);
            drawing.AddText(
                "Tilt",
                24,
                18,
                48,
                20,
                new OfficeFontInfo("Aptos", 14D),
                OfficeColor.Red,
                OfficeTextAlignment.Center,
                rotationDegrees: 35D,
                rotationCenterX: 48D,
                rotationCenterY: 28D);

            OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing);

            Assert.True(CountPixelsNear(image, OfficeColor.Red) > 20);
            Assert.True(AnyAlpha(image, 30, 10, 66, 46));
        }

        [Fact]
        public void OfficeDrawingRasterRenderer_RendersVerticallyAlignedTextThroughSharedTextRenderer() {
            OfficeDrawing drawing = new OfficeDrawing(96, 64);
            drawing.AddText(
                "Bottom",
                8,
                8,
                80,
                44,
                new OfficeFontInfo("Aptos", 12D),
                OfficeColor.Blue,
                OfficeTextAlignment.Right,
                verticalAlignment: OfficeTextVerticalAlignment.Bottom);

            OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing);
            var painted = Enumerable.Range(0, image.Width * image.Height)
                .Where(index => image.GetPixel(index % image.Width, index / image.Width).A > 0)
                .Select(index => (X: index % image.Width, Y: index / image.Width))
                .ToList();

            Assert.NotEmpty(painted);
            Assert.True(painted.Min(pixel => pixel.Y) > 28, "Expected bottom vertical alignment to place text in the lower part of the text box.");
            Assert.True(painted.Max(pixel => pixel.X) > 68, "Expected right horizontal alignment to place text near the right side of the text box.");
        }

        [Fact]
        public void OfficeDrawingRasterRenderer_FlattensBezierPathCommands() {
            OfficeDrawing drawing = new OfficeDrawing(64, 48);
            OfficeShape quadratic = OfficeShape.Path(
                OfficePathCommand.MoveTo(10, 40),
                OfficePathCommand.QuadraticBezierTo(30, 0, 50, 40));
            quadratic.StrokeColor = OfficeColor.Black;
            quadratic.StrokeWidth = 2;
            drawing.AddShape(quadratic, 0, 0);

            OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing);

            Assert.True(AnyAlpha(image, 28, 18, 32, 22));
            Assert.Equal(0, image.GetPixel(30, 40).A);
            Assert.True(CountPaintedPixels(image) > 60);
        }

        [Fact]
        public void OfficeRasterCanvas_DrawsPolylinesThroughMultipleSegments() {
            OfficeRasterImage image = new OfficeRasterImage(24, 24, OfficeColor.Transparent);
            OfficeRasterCanvas canvas = new OfficeRasterCanvas(image);

            canvas.DrawPolyline(new[] {
                new OfficePoint(3, 4),
                new OfficePoint(18, 4),
                new OfficePoint(18, 19)
            }, OfficeColor.Black, 1);

            Assert.True(image.GetPixel(8, 4).A > 0);
            Assert.True(image.GetPixel(18, 12).A > 0);
            Assert.Equal(0, image.GetPixel(5, 12).A);
        }

        [Fact]
        public void OfficeRasterCanvas_DrawsDashedPolylinesWithContinuousPatternAcrossSegments() {
            OfficeRasterImage continuous = new OfficeRasterImage(24, 11, OfficeColor.Transparent);
            OfficeRasterImage reset = new OfficeRasterImage(24, 11, OfficeColor.Transparent);
            var points = new[] {
                new OfficePoint(2, 5),
                new OfficePoint(10, 5),
                new OfficePoint(18, 5)
            };

            new OfficeRasterCanvas(continuous).DrawDashedPolyline(points, OfficeColor.Black, 1, dashLength: 6, gapLength: 4);
            new OfficeRasterCanvas(reset).DrawDashedPolyline(points, OfficeColor.Black, 1, dashLength: 6, gapLength: 4, resetDashPatternForEachSegment: true);

            Assert.True(continuous.GetPixel(4, 5).A > 0);
            Assert.Equal(0, continuous.GetPixel(10, 5).A);
            Assert.True(continuous.GetPixel(14, 5).A > 0);
            Assert.True(reset.GetPixel(10, 5).A > 0);
        }

        [Fact]
        public void OfficeRasterCanvas_DrawsStyledPolylinesAndEllipsesThroughSharedDashStyles() {
            OfficeRasterImage polyline = new OfficeRasterImage(42, 14, OfficeColor.Transparent);
            new OfficeRasterCanvas(polyline).DrawStyledPolyline(
                new[] {
                    new OfficePoint(2, 7),
                    new OfficePoint(20, 7),
                    new OfficePoint(38, 7)
                },
                OfficeColor.Black,
                2D,
                OfficeStrokeDashStyle.Dot);

            OfficeRasterImage ellipse = new OfficeRasterImage(48, 48, OfficeColor.Transparent);
            new OfficeRasterCanvas(ellipse).DrawStyledEllipse(
                24D,
                24D,
                18D,
                10D,
                OfficeColor.Transparent,
                OfficeColor.Black,
                2D,
                OfficeStrokeDashStyle.DashDot,
                rotationDegrees: 25D,
                rotationCenterX: 24D,
                rotationCenterY: 24D);

            Assert.True(AnyAlpha(polyline, 2, 6, 5, 8));
            Assert.True(CountTransparentColumnsOnRow(polyline, 7, 2, 38) >= 8);
            Assert.True(CountPaintedPixels(ellipse) > 20);
            Assert.Equal(0, ellipse.GetPixel(24, 24).A);
        }

        [Fact]
        public void OfficeRasterCanvas_RendersTuplePolygonsThroughSharedRasterPrimitives() {
            OfficeRasterImage filled = new OfficeRasterImage(22, 22, OfficeColor.Transparent);
            OfficeRasterCanvas fillCanvas = new OfficeRasterCanvas(filled);
            fillCanvas.FillPolygon(new[] {
                (2D, 2D),
                (19D, 2D),
                (11D, 18D)
            }, OfficeColor.Black);

            OfficeRasterImage evenOdd = new OfficeRasterImage(24, 24, OfficeColor.Transparent);
            OfficeRasterCanvas evenOddCanvas = new OfficeRasterCanvas(evenOdd);
            IReadOnlyList<(double X, double Y)>[] contours = {
                new[] {
                    (2D, 2D),
                    (21D, 2D),
                    (21D, 21D),
                    (2D, 21D)
                },
                new[] {
                    (8D, 8D),
                    (15D, 8D),
                    (15D, 15D),
                    (8D, 15D)
                }
            };
            evenOddCanvas.FillPolygonsEvenOdd(contours, OfficeColor.Black);

            OfficeRasterImage outline = new OfficeRasterImage(28, 24, OfficeColor.Transparent);
            new OfficeRasterCanvas(outline).DrawStyledPolygon(
                new[] {
                    (4D, 4D),
                    (23D, 4D),
                    (23D, 19D),
                    (4D, 19D)
                },
                OfficeColor.Black,
                2D,
                OfficeStrokeDashStyle.Dot,
                resetDashPatternForEachSegment: true);

            Assert.True(filled.GetPixel(11, 8).A > 0);
            Assert.True(evenOdd.GetPixel(4, 4).A > 0);
            Assert.Equal(0, evenOdd.GetPixel(11, 11).A);
            Assert.True(AnyAlpha(outline, 4, 3, 23, 5));
            Assert.True(AnyAlpha(outline, 3, 4, 5, 19));
            Assert.True(CountTransparentColumnsOnRow(outline, 4, 4, 23) >= 4);
        }

        private static bool AnyAlpha(OfficeRasterImage image, int left, int top, int right, int bottom) {
            for (int y = top; y <= bottom; y++) {
                for (int x = left; x <= right; x++) {
                    if (image.GetPixel(x, y).A > 0) {
                        return true;
                    }
                }
            }

            return false;
        }

        private static int CountTransparentColumnsOnRow(OfficeRasterImage image, int y, int left, int right) {
            int count = 0;
            for (int x = left; x <= right; x++) {
                if (image.GetPixel(x, y).A == 0) {
                    count++;
                }
            }

            return count;
        }

        [Fact]
        public void OfficeRasterCanvas_FillsMultipleContoursWithEvenOddRule() {
            OfficeRasterImage image = new OfficeRasterImage(22, 22, OfficeColor.Transparent);
            OfficeRasterCanvas canvas = new OfficeRasterCanvas(image);
            var contours = new List<IReadOnlyList<OfficePoint>> {
                new[] {
                    new OfficePoint(2, 2),
                    new OfficePoint(19, 2),
                    new OfficePoint(19, 19),
                    new OfficePoint(2, 19)
                },
                new[] {
                    new OfficePoint(7, 7),
                    new OfficePoint(14, 7),
                    new OfficePoint(14, 14),
                    new OfficePoint(7, 14)
                }
            };

            canvas.FillPolygonsEvenOdd(contours, OfficeColor.Black);

            Assert.True(image.GetPixel(4, 4).A > 0);
            Assert.Equal(0, image.GetPixel(10, 10).A);
        }

        [Fact]
        public void OfficeRasterCanvas_ScalesImagesWithInterpolation() {
            OfficeRasterImage source = new OfficeRasterImage(2, 1, OfficeColor.Transparent);
            source.SetPixel(0, 0, OfficeColor.Red);
            source.SetPixel(1, 0, OfficeColor.Blue);
            OfficeRasterImage target = new OfficeRasterImage(5, 1, OfficeColor.Transparent);
            OfficeRasterCanvas canvas = new OfficeRasterCanvas(target);

            canvas.DrawImage(source, 0, 0, 5, 1);

            OfficeColor middle = target.GetPixel(2, 0);
            Assert.True(middle.R > 40);
            Assert.True(middle.B > 40);
        }

        [Fact]
        public void OfficeRasterCanvas_DrawsRotatedImagesAroundCenter() {
            OfficeRasterImage source = new OfficeRasterImage(8, 2, OfficeColor.FromRgb(34, 197, 94));
            OfficeRasterImage target = new OfficeRasterImage(32, 32, OfficeColor.Transparent);
            OfficeRasterCanvas canvas = new OfficeRasterCanvas(target);

            canvas.DrawImage(source, 10, 14, 12, 4, rotationDegrees: 90, rotationCenterX: 16, rotationCenterY: 16);

            var painted = Enumerable.Range(0, target.Width * target.Height)
                .Where(index => target.GetPixel(index % target.Width, index / target.Width).A > 0)
                .Select(index => (X: index % target.Width, Y: index / target.Width))
                .ToList();
            Assert.NotEmpty(painted);
            Assert.True(painted.Max(pixel => pixel.Y) - painted.Min(pixel => pixel.Y) >
                        painted.Max(pixel => pixel.X) - painted.Min(pixel => pixel.X));
        }

        [Fact]
        public void OfficeRasterCanvas_FlipsCroppedImagesThroughSharedProjector() {
            OfficeRasterImage source = new OfficeRasterImage(6, 2, OfficeColor.Transparent);
            for (int y = 0; y < source.Height; y++) {
                source.SetPixel(0, y, OfficeColor.Red);
                source.SetPixel(1, y, OfficeColor.Red);
                source.SetPixel(2, y, OfficeColor.Blue);
                source.SetPixel(3, y, OfficeColor.Blue);
                source.SetPixel(4, y, OfficeColor.Lime);
                source.SetPixel(5, y, OfficeColor.Lime);
            }

            OfficeRasterImage target = new OfficeRasterImage(12, 4, OfficeColor.Transparent);
            OfficeRasterCanvas canvas = new OfficeRasterCanvas(target);

            canvas.DrawImage(
                source,
                2,
                1,
                8,
                2,
                sourceLeft: 2D / 6D,
                sourceTop: 0D,
                sourceWidth: 4D / 6D,
                sourceHeight: 1D,
                rotationDegrees: 0D,
                rotationCenterX: 6D,
                rotationCenterY: 2D,
                flipHorizontal: true,
                flipVertical: false);

            OfficeColor left = target.GetPixel(3, 1);
            OfficeColor right = target.GetPixel(8, 1);
            Assert.True(left.G > left.B, "Expected the horizontally flipped cropped image to place the green side on the left.");
            Assert.True(right.B > right.G, "Expected the horizontally flipped cropped image to place the blue side on the right.");
            Assert.Equal(0, target.GetPixel(1, 1).A);
        }

        [Fact]
        public void OfficeRasterCanvas_DrawsRotatedEllipsesAroundCenter() {
            OfficeRasterImage target = new OfficeRasterImage(48, 48, OfficeColor.Transparent);
            OfficeRasterCanvas canvas = new OfficeRasterCanvas(target);

            canvas.DrawEllipse(24, 24, 15, 4, OfficeColor.FromRgb(37, 99, 235), OfficeColor.Transparent, 0, rotationDegrees: 90, rotationCenterX: 24, rotationCenterY: 24);

            var painted = Enumerable.Range(0, target.Width * target.Height)
                .Where(index => target.GetPixel(index % target.Width, index / target.Width).A > 0)
                .Select(index => (X: index % target.Width, Y: index / target.Width))
                .ToList();
            Assert.NotEmpty(painted);
            Assert.True(painted.Max(pixel => pixel.Y) - painted.Min(pixel => pixel.Y) >
                        painted.Max(pixel => pixel.X) - painted.Min(pixel => pixel.X));
        }

        [Fact]
        public void OfficeRasterCanvas_DrawsEllipticalArcs() {
            OfficeRasterImage target = new OfficeRasterImage(48, 48, OfficeColor.Transparent);
            OfficeRasterCanvas canvas = new OfficeRasterCanvas(target);

            canvas.DrawArc(24, 24, 15, 10, 180, 360, OfficeColor.Black, 2);

            var painted = Enumerable.Range(0, target.Width * target.Height)
                .Where(index => target.GetPixel(index % target.Width, index / target.Width).A > 0)
                .Select(index => (X: index % target.Width, Y: index / target.Width))
                .ToList();
            Assert.NotEmpty(painted);
            Assert.Contains(painted, pixel => pixel.X < 14 && pixel.Y >= 21 && pixel.Y <= 27);
            Assert.Contains(painted, pixel => pixel.X > 34 && pixel.Y >= 21 && pixel.Y <= 27);
            Assert.True(painted.Count(pixel => pixel.Y < 24) > painted.Count(pixel => pixel.Y > 30));
        }

        [Fact]
        public void OfficeRasterCanvas_DrawsDashedEllipsesWithTransparentGaps() {
            OfficeRasterImage solid = new OfficeRasterImage(64, 64, OfficeColor.Transparent);
            OfficeRasterImage dashed = new OfficeRasterImage(64, 64, OfficeColor.Transparent);
            new OfficeRasterCanvas(solid).DrawEllipse(32, 32, 22, 12, OfficeColor.Transparent, OfficeColor.Black, 2, rotationDegrees: 35, rotationCenterX: 32, rotationCenterY: 32);
            new OfficeRasterCanvas(dashed).DrawDashedEllipse(32, 32, 22, 12, OfficeColor.Black, 2, dashLength: 5, gapLength: 5, rotationDegrees: 35, rotationCenterX: 32, rotationCenterY: 32);

            int solidPixels = CountPaintedPixels(solid);
            int dashedPixels = CountPaintedPixels(dashed);

            Assert.True(dashedPixels > solidPixels / 4, $"Expected visible dashed ellipse ink. dashed={dashedPixels}, solid={solidPixels}");
            Assert.True(dashedPixels < solidPixels * 85 / 100, $"Expected transparent gaps in dashed ellipse. dashed={dashedPixels}, solid={solidPixels}");
            Assert.Equal(0, dashed.GetPixel(32, 32).A);
        }

        [Fact]
        public void OfficeRasterCanvas_DrawsRotatedTextLines() {
            OfficeRasterImage target = new OfficeRasterImage(96, 96, OfficeColor.Transparent);
            OfficeRasterCanvas canvas = new OfficeRasterCanvas(target);

            canvas.DrawTextLine("TEXT", 48, 40, 18, OfficeColor.Black, bold: true, italic: true, OfficeTextAlignment.Center, rotationDegrees: 90, rotationCenterX: 48, rotationCenterY: 48);

            var painted = Enumerable.Range(0, target.Width * target.Height)
                .Where(index => target.GetPixel(index % target.Width, index / target.Width).A > 0)
                .Select(index => (X: index % target.Width, Y: index / target.Width))
                .ToList();
            Assert.NotEmpty(painted);
            Assert.True(painted.Max(pixel => pixel.Y) - painted.Min(pixel => pixel.Y) >
                        painted.Max(pixel => pixel.X) - painted.Min(pixel => pixel.X));
        }

        [Fact]
        public void OfficeRasterCanvas_PreservesGlyphCounters() {
            if (OfficeTrueTypeFont.TryLoadDefault() == null) {
                return;
            }

            OfficeRasterImage image = new OfficeRasterImage(80, 54, OfficeColor.Transparent);
            OfficeRasterCanvas canvas = new OfficeRasterCanvas(image);

            canvas.DrawText("O", 0, 0, 80, 54, OfficeColor.Black, 42, OfficeTextAlignment.Center, OfficeFontStyle.Regular);

            var inked = Enumerable.Range(0, image.Width * image.Height)
                .Where(index => image.GetPixel(index % image.Width, index / image.Width).A > 0)
                .Select(index => (X: index % image.Width, Y: index / image.Width))
                .ToList();
            Assert.NotEmpty(inked);
            int left = inked.Min(pixel => pixel.X);
            int right = inked.Max(pixel => pixel.X);
            int top = inked.Min(pixel => pixel.Y);
            int bottom = inked.Max(pixel => pixel.Y);

            bool hasTransparentCounter = false;
            for (int y = top + ((bottom - top) / 3); y <= bottom - ((bottom - top) / 3); y++) {
                for (int x = left + ((right - left) / 3); x <= right - ((right - left) / 3); x++) {
                    if (image.GetPixel(x, y).A == 0) {
                        hasTransparentCounter = true;
                    }
                }
            }

            Assert.True(hasTransparentCounter);
        }

        [Fact]
        public void OfficeTextBlockRenderer_DrawsRasterTextBlockWithAlignmentAndUnderline() {
            OfficeRasterImage image = new OfficeRasterImage(120, 72, OfficeColor.Transparent);
            OfficeRasterCanvas canvas = new OfficeRasterCanvas(image);
            OfficeTextBlockLayout layout = OfficeTextLayoutEngine.LayoutTextBlock(
                "Alpha\nBeta",
                12D,
                90D,
                48D,
                lineHeightFactor: 1.2D,
                minimumFontSize: 6D,
                canvas.MeasureText,
                wrap: true);

            OfficeTextBlockRenderer.DrawRasterTextBlock(
                canvas,
                layout,
                left: 10D,
                top: 10D,
                width: 90D,
                height: 48D,
                color: OfficeColor.Black,
                horizontalAlignment: OfficeTextAlignment.Right,
                verticalAlignment: OfficeTextVerticalAlignment.Bottom,
                bold: true,
                underline: true);

            var painted = Enumerable.Range(0, image.Width * image.Height)
                .Where(index => image.GetPixel(index % image.Width, index / image.Width).A > 0)
                .Select(index => (X: index % image.Width, Y: index / image.Width))
                .ToList();

            Assert.NotEmpty(painted);
            Assert.True(painted.Min(pixel => pixel.Y) >= 24, "Expected bottom vertical alignment to keep text in the lower part of the rectangle.");
            Assert.True(painted.Max(pixel => pixel.X) > 80, "Expected right horizontal alignment to place ink near the right side of the rectangle.");
            Assert.True(CountPaintedPixels(image) > 80);
        }

        [Fact]
        public void OfficeTextBlockRenderer_DrawsRasterRichTextBlockWithRunStyles() {
            OfficeRasterImage image = new OfficeRasterImage(140, 72, OfficeColor.Transparent);
            OfficeRasterCanvas canvas = new OfficeRasterCanvas(image);
            var line = new OfficeRichTextLine(new[] {
                new OfficeRichTextSegment("Red", canvas.MeasureText("Red", 14D), 14D, OfficeColor.Red, bold: true, italic: false, underline: true, fontFamily: "Aptos"),
                new OfficeRichTextSegment(" Blue", canvas.MeasureText(" Blue", 14D), 14D, OfficeColor.Blue, bold: false, italic: true, underline: false, fontFamily: "Aptos", strikethrough: true)
            });
            var layout = new OfficeRichTextBlockLayout(new[] { line }, lineHeight: 18D, width: line.Width, height: 18D);

            OfficeTextBlockRenderer.DrawRasterRichTextBlock(
                canvas,
                layout,
                left: 8D,
                top: 10D,
                width: 120D,
                height: 44D,
                horizontalAlignment: OfficeTextAlignment.Center,
                verticalAlignment: OfficeTextVerticalAlignment.Center);

            var painted = Enumerable.Range(0, image.Width * image.Height)
                .Where(index => image.GetPixel(index % image.Width, index / image.Width).A > 0)
                .Select(index => (X: index % image.Width, Y: index / image.Width))
                .ToList();

            Assert.NotEmpty(painted);
            Assert.True(CountPixelsNear(image, OfficeColor.Red) > 20, "Expected the first rich text run to render red ink.");
            Assert.True(CountPixelsNear(image, OfficeColor.Blue) > 20, "Expected the second rich text run to render blue ink.");
            Assert.True(painted.Min(pixel => pixel.X) > 20, "Expected centered rich text placement.");
            Assert.True(painted.Max(pixel => pixel.X) < 128, "Expected centered rich text placement to remain inside the target box.");
        }

        [Fact]
        public void OfficeTextBlockRenderer_AppendsSvgTextBlockWithSharedStyleAttributes() {
            var builder = new System.Text.StringBuilder();
            var layout = new OfficeTextBlockLayout(
                new[] {
                    new OfficeTextLine("A&B", 18D),
                    new OfficeTextLine("Beta", 24D)
                },
                fontSize: 10D,
                lineHeight: 12D,
                width: 24D,
                height: 24D);

            builder.AppendSvgTextBlock(
                layout,
                left: 0D,
                top: 0D,
                width: 100D,
                height: 40D,
                color: OfficeColor.FromRgba(1, 2, 3, 128),
                fontFamily: "Aptos, Arial, sans-serif",
                horizontalAlignment: OfficeTextAlignment.Center,
                verticalAlignment: OfficeTextVerticalAlignment.Bottom,
                bold: true,
                italic: true,
                underline: true,
                strikethrough: true,
                rotationDegrees: 15D,
                rotationCenterX: 50D,
                rotationCenterY: 20D);

            string svg = builder.ToString();
            Assert.Contains("text-anchor=\"middle\"", svg);
            Assert.Contains("font-family=\"Aptos, Arial, sans-serif\"", svg);
            Assert.Contains("font-weight=\"700\"", svg);
            Assert.Contains("font-style=\"italic\"", svg);
            Assert.Contains("text-decoration=\"underline line-through\"", svg);
            Assert.Contains("fill=\"#010203\"", svg);
            Assert.Contains("fill-opacity=\"0.502\"", svg);
            Assert.Contains("transform=\"rotate(15 50 20)\"", svg);
            Assert.Contains(">A&amp;B</text>", svg);
            Assert.Contains(">Beta</text>", svg);
        }

        [Fact]
        public void OfficeTextBlockRenderer_AppendsSvgTextElementWithTspans() {
            var builder = new System.Text.StringBuilder();

            builder.AppendSvgTextElement(
                "A&B\r\nBeta",
                x: 50D,
                y: 12D,
                lineHeight: 14D,
                color: OfficeColor.FromRgba(1, 2, 3, 128),
                fontFamily: "Aptos",
                fontSize: 10D,
                horizontalAlignment: OfficeTextAlignment.Center,
                bold: true,
                italic: true,
                underline: true,
                strikethrough: true,
                rotationDegrees: 15D,
                rotationCenterX: 50D,
                rotationCenterY: 20D);

            string svg = builder.ToString();
            Assert.Equal("<text x=\"50\" y=\"12\" font-family=\"Aptos\" font-size=\"10\" text-anchor=\"middle\" fill=\"#010203\" fill-opacity=\"0.502\" font-weight=\"700\" font-style=\"italic\" text-decoration=\"underline line-through\" transform=\"rotate(15 50 20)\">A&amp;B<tspan x=\"50\" dy=\"14\">Beta</tspan></text>", svg);
        }

        [Fact]
        public void OfficeTextBlockRenderer_AppendsSvgRichTextSegment() {
            var builder = new System.Text.StringBuilder();
            var segment = new OfficeRichTextSegment(
                "A&B",
                width: 24D,
                fontSize: 10D,
                color: OfficeColor.FromRgba(1, 2, 3, 128),
                bold: true,
                italic: true,
                underline: true,
                fontFamily: "Aptos",
                strikethrough: true);

            builder.AppendSvgRichTextSegment(segment, 5D, 12D);

            Assert.Equal("<text x=\"5\" y=\"12\" font-family=\"Aptos\" font-size=\"10\" text-anchor=\"start\" fill=\"#010203\" fill-opacity=\"0.502\" font-weight=\"700\" font-style=\"italic\" text-decoration=\"underline line-through\">A&amp;B</text>", builder.ToString());
        }

        [Fact]
        public void OfficeTextBlockRenderer_AppendsSvgRichTextBlockWithSharedPlacement() {
            var builder = new System.Text.StringBuilder();
            var firstLine = new OfficeRichTextLine(new[] {
                new OfficeRichTextSegment("Red", width: 20D, fontSize: 10D, color: OfficeColor.Red, bold: true, italic: false, underline: true, fontFamily: "Aptos"),
                new OfficeRichTextSegment("Blue", width: 30D, fontSize: 10D, color: OfficeColor.Blue, bold: false, italic: true, underline: false, fontFamily: "Aptos", strikethrough: true)
            });
            var secondLine = new OfficeRichTextLine(new[] {
                new OfficeRichTextSegment("Tail", width: 12D, fontSize: 8D, color: OfficeColor.FromRgb(1, 2, 3), bold: false, italic: false, underline: false, fontFamily: "Aptos")
            });
            var layout = new OfficeRichTextBlockLayout(
                new[] { firstLine, secondLine },
                lineHeight: 14D,
                width: firstLine.Width,
                height: 28D);

            builder.AppendSvgRichTextBlock(
                layout,
                left: 0D,
                top: 0D,
                width: 100D,
                height: 50D,
                OfficeTextAlignment.Center,
                OfficeTextVerticalAlignment.Bottom,
                rotationDegrees: 15D,
                rotationCenterX: 50D,
                rotationCenterY: 25D);

            string svg = builder.ToString();
            Assert.Contains("x=\"25\" y=\"32.4\"", svg);
            Assert.Contains("x=\"45\" y=\"32.4\"", svg);
            Assert.Contains("x=\"44\" y=\"45.72\"", svg);
            Assert.Contains("font-weight=\"700\"", svg);
            Assert.Contains("font-style=\"italic\"", svg);
            Assert.Contains("text-decoration=\"line-through\"", svg);
            Assert.Contains("transform=\"rotate(15 50 25)\"", svg);
            Assert.Contains(">Red</text>", svg);
            Assert.Contains(">Blue</text>", svg);
            Assert.Contains(">Tail</text>", svg);
        }

        [Fact]
        public void OfficeTextBlockRenderer_WritesSvgTextBlockWithTspansAndAdapterAttributes() {
            var layout = new OfficeTextBlockLayout(
                new[] {
                    new OfficeTextLine("A&B", 18D),
                    new OfficeTextLine("Beta", 24D)
                },
                fontSize: 10D,
                lineHeight: 12D,
                width: 24D,
                height: 24D);
            var builder = new System.Text.StringBuilder();
            using (var writer = System.Xml.XmlWriter.Create(
                builder,
                new System.Xml.XmlWriterSettings {
                    ConformanceLevel = System.Xml.ConformanceLevel.Fragment,
                    OmitXmlDeclaration = true
                })) {
                OfficeTextBlockRenderer.WriteSvgTextBlock(
                    writer,
                    layout,
                    left: 0D,
                    top: 0D,
                    width: 100D,
                    height: 40D,
                    color: OfficeColor.FromRgba(1, 2, 3, 128),
                    fontFamily: "Aptos, Arial, sans-serif",
                    horizontalAlignment: OfficeTextAlignment.Center,
                    verticalAlignment: OfficeTextVerticalAlignment.Bottom,
                    bold: true,
                    italic: true,
                    underline: true,
                    rotationDegrees: 15D,
                    rotationCenterX: 50D,
                    rotationCenterY: 20D,
                    svgNamespace: "http://www.w3.org/2000/svg",
                    configureTextAttributes: textWriter => textWriter.WriteAttributeString("data-officeimo-test", "true"),
                    strikethrough: true);
            }

            string svg = builder.ToString();
            Assert.Contains("<text", svg);
            Assert.Contains("data-officeimo-test=\"true\"", svg);
            Assert.Contains("text-anchor=\"middle\"", svg);
            Assert.Contains("dominant-baseline=\"middle\"", svg);
            Assert.Contains("font-family=\"Aptos, Arial, sans-serif\"", svg);
            Assert.Contains("font-weight=\"700\"", svg);
            Assert.Contains("font-style=\"italic\"", svg);
            Assert.Contains("text-decoration=\"underline line-through\"", svg);
            Assert.Contains("fill=\"#010203\"", svg);
            Assert.Contains("fill-opacity=\"0.502\"", svg);
            Assert.Contains("transform=\"rotate(15 50 20)\"", svg);
            Assert.Contains("<tspan", svg);
            Assert.Contains(">A&amp;B</tspan>", svg);
            Assert.Contains(">Beta</tspan>", svg);
        }

        private static int CountPaintedPixels(OfficeRasterImage image) =>
            Enumerable.Range(0, image.Width * image.Height)
                .Count(index => image.GetPixel(index % image.Width, index / image.Width).A > 0);

        private static int CountPixelsNear(OfficeRasterImage image, OfficeColor expected) =>
            Enumerable.Range(0, image.Width * image.Height)
                .Count(index => {
                    OfficeColor color = image.GetPixel(index % image.Width, index / image.Width);
                    return Math.Abs(color.R - expected.R) <= 8 &&
                        Math.Abs(color.G - expected.G) <= 8 &&
                        Math.Abs(color.B - expected.B) <= 8 &&
                        color.A > 0;
                });

        private static void AssertColorNear(OfficeColor actual, OfficeColor expected) {
            Assert.True(
                Math.Abs(actual.R - expected.R) <= 8 &&
                Math.Abs(actual.G - expected.G) <= 8 &&
                Math.Abs(actual.B - expected.B) <= 8 &&
                actual.A > 0,
                $"Expected ARGB near {expected.A},{expected.R},{expected.G},{expected.B} but got {actual.A},{actual.R},{actual.G},{actual.B}.");
        }

    }
}
