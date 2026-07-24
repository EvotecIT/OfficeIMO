using OfficeIMO.Drawing;
using System.Collections;
using System.Collections.Generic;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class DrawingRasterTests {
        [Fact]
        public void OfficeRasterCanvas_TransparentTuplePathsDoNotMaterializePoints() {
            OfficeRasterImage image = new OfficeRasterImage(2, 2, OfficeColor.Transparent);
            OfficeRasterCanvas canvas = new OfficeRasterCanvas(image);
            var points = new ExplosiveTuplePointList();

            canvas.DrawPolyline(points, OfficeColor.Transparent, 1D);
            canvas.DrawStyledPolyline(points, OfficeColor.Transparent, 1D);
            canvas.DrawPatternedPolyline(points, OfficeColor.Transparent, 1D, new[] { 1D, 1D });
            canvas.DrawDashedPolyline(points, OfficeColor.Transparent, 1D);
            canvas.DrawStyledPolygon(points, OfficeColor.Transparent, 1D);
            canvas.FillPolygon(points, OfficeColor.Transparent);
        }

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
        public void OfficeRasterImageDecoder_DecodesUncompressedBmp24RowsThroughSharedRasterPath() {
            byte[] bmp = CreateBmp24(
                2,
                2,
                new[] {
                    OfficeColor.Red, OfficeColor.Lime,
                    OfficeColor.Blue, OfficeColor.White
                });

            Assert.True(OfficeRasterImageDecoder.TryDecode(bmp, out OfficeRasterImage? image));
            Assert.Equal(2, image!.Width);
            Assert.Equal(2, image.Height);
            Assert.Equal(OfficeColor.Red, image.GetPixel(0, 0));
            Assert.Equal(OfficeColor.Lime, image.GetPixel(1, 0));
            Assert.Equal(OfficeColor.Blue, image.GetPixel(0, 1));
            Assert.Equal(OfficeColor.White, image.GetPixel(1, 1));
        }

        [Fact]
        public void OfficeDibReader_And_PngConverter_Decode_RtfStyle_DibPayload() {
            byte[] bmp = CreateBmp24(2, 1, new[] { OfficeColor.Red, OfficeColor.Blue });
            byte[] dib = bmp.Skip(14).ToArray();

            Assert.True(OfficeDibReader.TryDecode(dib, out OfficeRasterImage? image));
            Assert.Equal(OfficeColor.Red, image!.GetPixel(0, 0));
            Assert.Equal(OfficeColor.Blue, image.GetPixel(1, 0));
            Assert.True(OfficeImagePngConverter.TryConvertDibToPng(dib, out byte[] png));
            Assert.True(OfficePngReader.TryDecode(png, out OfficeRasterImage? roundTrip));
            Assert.Equal(OfficeColor.Red, roundTrip!.GetPixel(0, 0));
            Assert.Equal(OfficeColor.Blue, roundTrip.GetPixel(1, 0));
        }

        [Fact]
        public void OfficeImagePngConverter_RejectsNonzeroFrameForDibFallback() {
            byte[] bmp = CreateBmp24(1, 1, new[] { OfficeColor.Red });
            byte[] dib = bmp.Skip(14).ToArray();
            var options = new OfficeRasterDecodeOptions { FrameIndex = 1 };

            Assert.False(OfficeImagePngConverter.TryConvertToPng(
                dib,
                options,
                out byte[] png,
                out OfficeRasterDecodeInfo decodeInfo));
            Assert.Empty(png);
            Assert.False(decodeInfo.Succeeded);
            Assert.Equal(1, decodeInfo.SelectedFrameIndex);
        }

        [Fact]
        public void OfficeRasterImageDecoder_DecodesTopDownBmp24RowsThroughSharedRasterPath() {
            byte[] bmp = CreateBmp24(
                2,
                2,
                new[] {
                    OfficeColor.Red, OfficeColor.Lime,
                    OfficeColor.Blue, OfficeColor.White
                },
                topDown: true);

            Assert.True(OfficeRasterImageDecoder.TryDecode(bmp, out OfficeRasterImage? image));
            Assert.Equal(2, image!.Width);
            Assert.Equal(2, image.Height);
            Assert.Equal(OfficeColor.Red, image.GetPixel(0, 0));
            Assert.Equal(OfficeColor.Lime, image.GetPixel(1, 0));
            Assert.Equal(OfficeColor.Blue, image.GetPixel(0, 1));
            Assert.Equal(OfficeColor.White, image.GetPixel(1, 1));
        }

        [Fact]
        public void OfficeRasterImageDecoder_DecodesUncompressedBmp32AlphaThroughSharedRasterPath() {
            byte[] bmp = CreateBmp32(
                2,
                2,
                new[] {
                    OfficeColor.FromRgba(255, 0, 0, 255), OfficeColor.FromRgba(0, 255, 0, 192),
                    OfficeColor.FromRgba(0, 0, 255, 128), OfficeColor.FromRgba(255, 255, 255, 64)
                });

            Assert.True(OfficeRasterImageDecoder.TryDecode(bmp, out OfficeRasterImage? image));
            Assert.Equal(2, image!.Width);
            Assert.Equal(2, image.Height);
            Assert.Equal(OfficeColor.FromRgba(255, 0, 0, 255), image.GetPixel(0, 0));
            Assert.Equal(OfficeColor.FromRgba(0, 255, 0, 192), image.GetPixel(1, 0));
            Assert.Equal(OfficeColor.FromRgba(0, 0, 255, 128), image.GetPixel(0, 1));
            Assert.Equal(OfficeColor.FromRgba(255, 255, 255, 64), image.GetPixel(1, 1));
        }

        [Fact]
        public void OfficeRasterImageDecoder_DecodesTopDownBmp32AlphaThroughSharedRasterPath() {
            byte[] bmp = CreateBmp32(
                2,
                2,
                new[] {
                    OfficeColor.FromRgba(255, 0, 0, 255), OfficeColor.FromRgba(0, 255, 0, 192),
                    OfficeColor.FromRgba(0, 0, 255, 128), OfficeColor.FromRgba(255, 255, 255, 64)
                },
                topDown: true);

            Assert.True(OfficeRasterImageDecoder.TryDecode(bmp, out OfficeRasterImage? image));
            Assert.Equal(2, image!.Width);
            Assert.Equal(2, image.Height);
            Assert.Equal(OfficeColor.FromRgba(255, 0, 0, 255), image.GetPixel(0, 0));
            Assert.Equal(OfficeColor.FromRgba(0, 255, 0, 192), image.GetPixel(1, 0));
            Assert.Equal(OfficeColor.FromRgba(0, 0, 255, 128), image.GetPixel(0, 1));
            Assert.Equal(OfficeColor.FromRgba(255, 255, 255, 64), image.GetPixel(1, 1));
        }

        [Fact]
        public void OfficeRasterImageDecoder_TreatsBmp32ReservedBytesAsOpaqueWhenAlphaIsAbsent() {
            byte[] bmp = CreateBmp32(
                2,
                2,
                new[] {
                    OfficeColor.FromRgba(255, 0, 0, 0), OfficeColor.FromRgba(0, 255, 0, 0),
                    OfficeColor.FromRgba(0, 0, 255, 0), OfficeColor.FromRgba(255, 255, 255, 0)
                });

            Assert.True(OfficeRasterImageDecoder.TryDecode(bmp, out OfficeRasterImage? image));
            Assert.Equal(OfficeColor.Red, image!.GetPixel(0, 0));
            Assert.Equal(OfficeColor.Lime, image.GetPixel(1, 0));
            Assert.Equal(OfficeColor.Blue, image.GetPixel(0, 1));
            Assert.Equal(OfficeColor.White, image.GetPixel(1, 1));
        }

        [Fact]
        public void OfficeDrawingRasterRenderer_PaintsDecodedBmpImages() {
            byte[] bmp = CreateBmp24(1, 1, new[] { OfficeColor.FromRgb(18, 52, 86) });
            OfficeDrawing drawing = new OfficeDrawing(20, 16);
            drawing.AddImage(
                bmp,
                "image/bmp",
                new OfficeImageProjection(new OfficeImagePlacement(4, 3, 8, 6)),
                "BMP marker");

            OfficeRasterImage rendered = OfficeDrawingRasterRenderer.Render(drawing);

            Assert.Equal(OfficeColor.FromRgb(18, 52, 86), rendered.GetPixel(7, 5));
        }

        [Fact]
        public void OfficeDrawingRasterRenderer_BlendsDecodedBmp32AlphaImages() {
            byte[] bmp = CreateBmp32(1, 1, new[] { OfficeColor.FromRgba(255, 0, 0, 128) });
            OfficeDrawing drawing = new OfficeDrawing(20, 16);
            drawing.AddImage(
                bmp,
                "image/bmp",
                new OfficeImageProjection(new OfficeImagePlacement(4, 3, 8, 6)),
                "BMP alpha marker");

            OfficeRasterImage rendered = OfficeDrawingRasterRenderer.Render(drawing, background: OfficeColor.White);
            OfficeColor blended = rendered.GetPixel(7, 5);

            Assert.True(blended.R >= 252, $"Expected red channel to stay near full after alpha blend, got {blended.R}.");
            Assert.InRange(blended.G, 124, 130);
            Assert.InRange(blended.B, 124, 130);
            Assert.Equal(255, blended.A);
        }

        [Fact]
        public void OfficeDrawingRasterRenderer_RasterizesSupportedSvgImagesAtDestinationResolution() {
            byte[] svg = System.Text.Encoding.UTF8.GetBytes(
                "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 10 10\">" +
                "<rect x=\"0\" y=\"0\" width=\"10\" height=\"10\" fill=\"#D7263D\"/>" +
                "</svg>");
            var diagnostics = new List<OfficeImageExportDiagnostic>();
            var fallback = new OfficeRasterImageFallbackCodec(
                diagnostics: diagnostics,
                source: "SVG media");
            var drawing = new OfficeDrawing(80, 60);
            drawing.AddImage(
                svg,
                "image/svg+xml",
                new OfficeImageProjection(new OfficeImagePlacement(10, 10, 40, 30)),
                "SVG marker");

            OfficeRasterImage rendered = OfficeDrawingRasterRenderer.Render(
                drawing,
                new OfficeDrawingRasterRenderOptions {
                    ImageCodec = fallback
                });

            Assert.Equal(OfficeColor.FromRgb(0xD7, 0x26, 0x3D), rendered.GetPixel(30, 25));
            Assert.Empty(diagnostics);
        }

        [Fact]
        public void OfficeDrawingRasterRenderer_KeepsUnsupportedSvgMediaOnTheVisibleFallbackPath() {
            byte[] svg = System.Text.Encoding.UTF8.GetBytes(
                "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 10 10\">" +
                "<foreignObject x=\"0\" y=\"0\" width=\"10\" height=\"10\"/>" +
                "</svg>");
            var diagnostics = new List<OfficeImageExportDiagnostic>();
            var fallback = new OfficeRasterImageFallbackCodec(
                diagnostics: diagnostics,
                source: "SVG media");
            var drawing = new OfficeDrawing(40, 40);
            drawing.AddImage(
                svg,
                "image/svg+xml",
                new OfficeImageProjection(new OfficeImagePlacement(4, 4, 32, 32)),
                "Unsupported SVG marker");

            OfficeDrawingRasterRenderer.Render(
                drawing,
                new OfficeDrawingRasterRenderOptions {
                    ImageCodec = fallback
                });

            OfficeImageExportDiagnostic diagnostic = Assert.Single(diagnostics);
            Assert.Equal(OfficeImageExportDiagnosticCodes.SourceImageDecodeFallback, diagnostic.Code);
            Assert.Equal(OfficeImageExportLossKind.Omission, diagnostic.LossKind);
        }

        [Fact]
        public void OfficeSvgImageRenderer_CreatesPngDataUriForDecodedBmpImages() {
            byte[] bmp = CreateBmp24(1, 1, new[] { OfficeColor.FromRgb(18, 52, 86) });

            Assert.False(OfficeSvgImageRenderer.TryResolveEmbeddableContentType("image/bmp", bmp, null, out string directContentType));
            Assert.Equal(string.Empty, directContentType);
            Assert.True(OfficeSvgImageRenderer.TryCreateDataUri("image/bmp", bmp, null, out string dataUri));

            const string prefix = "data:image/png;base64,";
            Assert.StartsWith(prefix, dataUri, StringComparison.Ordinal);
            byte[] png = Convert.FromBase64String(dataUri.Substring(prefix.Length));
            Assert.True(OfficePngReader.TryDecode(png, out OfficeRasterImage? decoded));
            Assert.Equal(OfficeColor.FromRgb(18, 52, 86), decoded!.GetPixel(0, 0));
        }

        [Fact]
        public void OfficeDrawingSvgExporter_EmbedsDecodedBmpImagesAsPngDataUris() {
            byte[] bmp = CreateBmp24(1, 1, new[] { OfficeColor.FromRgb(18, 52, 86) });
            OfficeDrawing drawing = new OfficeDrawing(20, 16);
            drawing.AddImage(
                bmp,
                "image/bmp",
                new OfficeImageProjection(new OfficeImagePlacement(4, 3, 8, 6)),
                "BMP marker");

            string svg = OfficeDrawingSvgExporter.ToSvg(drawing);

            Assert.Contains("<image", svg, StringComparison.Ordinal);
            Assert.Contains("data:image/png;base64,", svg, StringComparison.Ordinal);
            Assert.DoesNotContain("data:image/bmp", svg, StringComparison.Ordinal);
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
        public void OfficeRasterCanvas_FillsMultiStopLinearGradientRectangle() {
            OfficeRasterImage image = new OfficeRasterImage(41, 10, OfficeColor.Transparent);
            OfficeRasterCanvas canvas = new OfficeRasterCanvas(image);
            var gradient = new OfficeLinearGradient(
                0,
                0.5,
                1,
                0.5,
                new[] {
                    new OfficeGradientStop(0D, OfficeColor.Red),
                    new OfficeGradientStop(0.5D, OfficeColor.Lime),
                    new OfficeGradientStop(1D, OfficeColor.Blue)
                });

            canvas.FillLinearGradientRectangle(0, 0, image.Width, image.Height, gradient);

            OfficeColor left = image.GetPixel(1, 5);
            OfficeColor middle = image.GetPixel(20, 5);
            OfficeColor right = image.GetPixel(39, 5);
            Assert.True(left.R > left.G && left.R > left.B, $"Expected left gradient edge to be red-ish, got {left}.");
            Assert.True(middle.G > 220 && middle.R < 40 && middle.B < 40, $"Expected middle gradient stop to be green, got {middle}.");
            Assert.True(right.B > right.R && right.B > right.G, $"Expected right gradient edge to be blue-ish, got {right}.");
        }

        [Fact]
        public void OfficeRasterCanvas_PreservesDuplicateOffsetHardStops() {
            OfficeRasterImage image = new OfficeRasterImage(100, 10, OfficeColor.Transparent);
            OfficeRasterCanvas canvas = new OfficeRasterCanvas(image);
            var gradient = new OfficeLinearGradient(
                0D,
                0.5D,
                1D,
                0.5D,
                new[] {
                    new OfficeGradientStop(0D, OfficeColor.Red),
                    new OfficeGradientStop(0.5D, OfficeColor.Red),
                    new OfficeGradientStop(0.5D, OfficeColor.Blue),
                    new OfficeGradientStop(1D, OfficeColor.Blue)
                });

            canvas.FillLinearGradientRectangle(0D, 0D, image.Width, image.Height, gradient);

            Assert.Equal(OfficeColor.Red, image.GetPixel(48, 5));
            Assert.Equal(OfficeColor.Blue, image.GetPixel(52, 5));
        }

        [Fact]
        public void OfficeRasterCanvas_FillsOffCenterEllipticalRadialGradient() {
            OfficeRasterImage image = new OfficeRasterImage(100, 100, OfficeColor.Transparent);
            OfficeRasterCanvas canvas = new OfficeRasterCanvas(image);
            var gradient = new OfficeRadialGradient(
                0.25D,
                0.5D,
                0D,
                0D,
                0.25D,
                0.5D,
                0.5D,
                0.25D,
                new[] {
                    new OfficeGradientStop(0D, OfficeColor.Red),
                    new OfficeGradientStop(1D, OfficeColor.Blue)
                });

            canvas.FillRadialGradientRectangle(0D, 0D, image.Width, image.Height, gradient);

            OfficeColor center = image.GetPixel(24, 49);
            OfficeColor horizontalEdge = image.GetPixel(74, 49);
            OfficeColor verticalEdge = image.GetPixel(24, 74);
            Assert.True(center.R > center.B, $"Expected the authored center to be red-ish, got {center}.");
            Assert.True(horizontalEdge.B > horizontalEdge.R, $"Expected the horizontal ellipse edge to be blue-ish, got {horizontalEdge}.");
            Assert.True(verticalEdge.B > verticalEdge.R, $"Expected the vertical ellipse edge to be blue-ish, got {verticalEdge}.");
        }

        [Fact]
        public void OfficeRasterCanvas_IgnoresRectanglesFullyOutsideCanvas() {
            OfficeRasterImage image = new OfficeRasterImage(6, 6, OfficeColor.Transparent);
            OfficeRasterCanvas canvas = new OfficeRasterCanvas(image);

            canvas.FillRectangle(8, 1, 4, 4, OfficeColor.Red);
            canvas.FillRectangle(1, 8, 4, 4, OfficeColor.Blue);
            canvas.FillLinearGradientRectangle(8, 8, 4, 4, OfficeLinearGradient.Horizontal(OfficeColor.Red, OfficeColor.Blue));

            Assert.Equal(0, CountPaintedPixels(image));
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
        public void OfficeRasterCanvas_DrawsSharedPercentStipplePatterns() {
            OfficeRasterImage sparse = new OfficeRasterImage(16, 16, OfficeColor.Transparent);
            OfficeRasterImage denser = new OfficeRasterImage(16, 16, OfficeColor.Transparent);

            new OfficeRasterCanvas(sparse).DrawHatchPatternRectangle(0, 0, 16, 16, OfficeColor.Green, 4, 2, OfficeHatchPatternKind.Percent6_25);
            new OfficeRasterCanvas(denser).DrawHatchPatternRectangle(0, 0, 16, 16, OfficeColor.Green, 4, 2, OfficeHatchPatternKind.Percent12_5);

            int sparsePixels = CountPaintedPixels(sparse);
            int denserPixels = CountPaintedPixels(denser);
            Assert.Equal(16, sparsePixels);
            Assert.Equal(32, denserPixels);
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
        public void OfficeSparklineRenderer_UsesExplicitScaleDomainForColumnHeights() {
            var autoBuilder = new System.Text.StringBuilder();
            OfficeSparklineRenderer.AppendSvg(
                autoBuilder,
                0,
                0,
                80,
                28,
                new[] { 5D },
                OfficeSparklineKind.Column,
                new OfficeSparklineStyle { Padding = 0D });

            var scaledBuilder = new System.Text.StringBuilder();
            OfficeSparklineRenderer.AppendSvg(
                scaledBuilder,
                0,
                0,
                80,
                28,
                new[] { 5D },
                OfficeSparklineKind.Column,
                new OfficeSparklineStyle {
                    Padding = 0D,
                    MinimumValue = 0D,
                    MaximumValue = 10D
                });

            double autoHeight = ExtractSvgRectHeight(autoBuilder.ToString());
            double scaledHeight = ExtractSvgRectHeight(scaledBuilder.ToString());
            Assert.True(autoHeight > scaledHeight * 1.8D, $"Expected explicit scale domain to reduce bar height. Auto={autoHeight}; Scaled={scaledHeight}.");
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
        public void OfficeConditionalIconRenderer_DrawsReusableRasterIcons() {
            OfficeRasterImage circle = new OfficeRasterImage(24, 24, OfficeColor.Transparent);
            OfficeRasterImage arrow = new OfficeRasterImage(24, 24, OfficeColor.Transparent);
            OfficeRasterImage rating = new OfficeRasterImage(24, 24, OfficeColor.Transparent);
            OfficeRasterImage quarter = new OfficeRasterImage(24, 24, OfficeColor.Transparent);
            OfficeRasterImage flag = new OfficeRasterImage(24, 24, OfficeColor.Transparent);

            OfficeConditionalIconRenderer.DrawRaster(new OfficeRasterCanvas(circle), 3, 3, 18, OfficeConditionalIconKind.RedCircle, scale: 1D);
            OfficeConditionalIconRenderer.DrawRaster(new OfficeRasterCanvas(arrow), 3, 3, 18, OfficeConditionalIconKind.GreenUpArrow, scale: 1D);
            OfficeConditionalIconRenderer.DrawRaster(new OfficeRasterCanvas(rating), 3, 3, 18, OfficeConditionalIconKind.RatingFive, scale: 1D);
            OfficeConditionalIconRenderer.DrawRaster(new OfficeRasterCanvas(quarter), 3, 3, 18, OfficeConditionalIconKind.QuarterOne, scale: 1D);
            OfficeConditionalIconRenderer.DrawRaster(new OfficeRasterCanvas(flag), 3, 3, 18, OfficeConditionalIconKind.GreenFlag, scale: 1D);

            Assert.True(CountPixelsNear(circle, OfficeColor.FromRgb(220, 38, 38)) > 40);
            Assert.True(CountPixelsNear(arrow, OfficeColor.FromRgb(22, 163, 74)) > 30);
            Assert.True(CountPixelsNear(rating, OfficeColor.FromRgb(22, 163, 74)) > 30);
            Assert.True(CountPixelsNear(quarter, OfficeColor.FromRgb(249, 115, 22)) > 20);
            Assert.True(CountPixelsNear(flag, OfficeColor.FromRgb(22, 163, 74)) > 25);
            Assert.True(CountPixelsNearAlpha(circle, OfficeColor.FromRgb(15, 23, 42), 8, 10, 70) > 0);
            Assert.True(CountPixelsNearAlpha(arrow, OfficeColor.FromRgb(15, 23, 42), 8, 10, 70) > 0);
            Assert.True(CountPixelsNearAlpha(rating, OfficeColor.FromRgb(15, 23, 42), 8, 10, 70) > 0);
            Assert.Equal(0, circle.GetPixel(0, 0).A);
            Assert.Equal(0, arrow.GetPixel(0, 0).A);
            Assert.Equal(0, rating.GetPixel(0, 0).A);
            Assert.Equal(0, quarter.GetPixel(0, 0).A);
            Assert.Equal(0, flag.GetPixel(0, 0).A);
        }

        [Fact]
        public void OfficeConditionalIconRenderer_AppendsReusableRatingAndQuarterSvgIcons() {
            var builder = new System.Text.StringBuilder();

            OfficeConditionalIconRenderer.AppendSvg(builder, 2, 3, 18, OfficeConditionalIconKind.RatingThree, scale: 1D);
            OfficeConditionalIconRenderer.AppendSvg(builder, 24, 3, 18, OfficeConditionalIconKind.QuarterTwo, scale: 1D);
            OfficeConditionalIconRenderer.AppendSvg(builder, 46, 3, 18, OfficeConditionalIconKind.GreenFlag, scale: 1D);
            string svg = builder.ToString();

            Assert.DoesNotContain("<rect", svg, StringComparison.Ordinal);
            Assert.True(CountOccurrences(svg, "<polygon") >= 6, svg);
            Assert.Contains("#F59E0B", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#16A34A", svg, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void OfficeCalloutRenderer_DrawsSharedRasterAndSvgCallouts() {
            var callout = new OfficeCallout(
                x: 36D,
                y: 18D,
                width: 132D,
                height: 74D,
                anchorX: 20D,
                anchorY: 34D,
                title: "Reviewer",
                text: "Ready for leadership review");
            var style = new OfficeCalloutStyle {
                AccentColor = OfficeColor.FromRgb(124, 58, 237)
            };
            OfficeRasterImage image = new OfficeRasterImage(220, 120, OfficeColor.Transparent);
            OfficeRasterCanvas canvas = new OfficeRasterCanvas(image);

            OfficeCalloutRenderer.DrawRaster(canvas, callout, style);

            Assert.True(CountPixelsNear(image, style.FillColor) > 1000);
            Assert.True(CountPixelsNear(image, style.HeaderFillColor) > 400);
            Assert.True(CountPixelsNear(image, style.AccentColor) > 120);
            Assert.True(CountPixelsNearAlpha(image, style.ShadowColor, 8, 10, 80) > 100);
            Assert.True(CountPaintedPixels(image) > 3000);

            var builder = new System.Text.StringBuilder();
            OfficeCalloutRenderer.AppendSvg(
                builder,
                callout,
                style,
                canvas.MeasureText,
                idPrefix: "review-callout");
            string svg = builder.ToString();

            Assert.Contains("review-callout-body", svg, StringComparison.Ordinal);
            Assert.Contains("review-callout-text", svg, StringComparison.Ordinal);
            Assert.Contains("fill-opacity=", svg, StringComparison.Ordinal);
            Assert.Contains("fill=\"#FFFBE6\"", svg, StringComparison.Ordinal);
            Assert.Contains("fill=\"#0F172A\"", svg, StringComparison.Ordinal);
            Assert.Contains("fill=\"#7C3AED\"", svg, StringComparison.Ordinal);
            Assert.Contains(">Reviewer</text>", svg, StringComparison.Ordinal);
            Assert.Contains("Ready", svg, StringComparison.Ordinal);
            Assert.Contains("leadership", svg, StringComparison.Ordinal);
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
        public void OfficeTextLayoutEngine_PreservesWhitespaceRunsWhenWrappedTextStillFits() {
            static double Measure(string? value, double size) => (value?.Length ?? 0) * size;

            IReadOnlyList<OfficeTextLine> lines = OfficeTextLayoutEngine.WrapLines("  A  B\tC", 1D, 20D, Measure);

            OfficeTextLine line = Assert.Single(lines);
            Assert.Equal("  A  B  C", line.Text);
            Assert.Equal(9D, line.Width);
        }

        [Fact]
        public void OfficeTextLayoutEngine_DropsSoftWrapWhitespaceInsteadOfEmittingEmptyLine() {
            static double Measure(string? value, double size) => (value?.Length ?? 0) * size;

            IReadOnlyList<OfficeTextLine> lines = OfficeTextLayoutEngine.WrapLines("No old unsupported diagnostic", 1D, 10D, Measure);

            Assert.DoesNotContain(lines, line => line.Text.Length == 0);
            Assert.Equal("diagnostic", lines[lines.Count - 1].Text);
        }

        [Fact]
        public void OfficeTextLayoutEngine_ExpandsTabsThroughPlainAndRichTextLayout() {
            static double Measure(string? value, double size) => (value?.Length ?? 0) * size;

            IReadOnlyList<OfficeTextLine> lines = OfficeTextLayoutEngine.WrapLines("A\tB", 10D, 100D, Measure);
            OfficeTextLine singleLine = OfficeTextLayoutEngine.TrimLineToWidth("A\tB", 10D, 100D, Measure, out bool clipped);
            OfficeRichTextBlockLayout rich = OfficeTextLayoutEngine.LayoutRichTextBlock(
                new[] { new OfficeRichTextRun("A\tB", 10D, OfficeColor.Black) },
                100D,
                30D,
                lineHeightFactor: 1.2D,
                (value, size, _) => Measure(value, size),
                wrap: true);

            Assert.False(clipped);
            Assert.Equal("A   B", lines.Single().Text);
            Assert.Equal(50D, lines.Single().Width);
            Assert.Equal("A   B", singleLine.Text);
            Assert.Equal(50D, singleLine.Width);
            Assert.Equal("A   B", string.Concat(rich.Lines.Single().Segments.Select(segment => segment.Text)));
            Assert.Equal(50D, rich.Lines.Single().Width);
        }

        [Fact]
        public void OfficeTextLayoutEngine_ProjectsParagraphIndentThroughPlainAndRichTextLines() {
            static double Measure(string? value, double size) => (value?.Length ?? 0) * size;

            OfficeTextParagraphIndent hanging = OfficeTextParagraphIndent.Hanging(2D);
            OfficeTextBlockLayout plain = OfficeTextLayoutEngine.LayoutTextBlock(
                "Alpha Beta Gamma",
                1D,
                8D,
                50D,
                1.2D,
                1D,
                Measure,
                wrap: true,
                paragraphIndent: hanging);
            OfficeRichTextBlockLayout rich = OfficeTextLayoutEngine.LayoutRichTextBlock(
                new[] { new OfficeRichTextRun("Alpha Beta Gamma", 1D, OfficeColor.Black) },
                8D,
                50D,
                1.2D,
                (value, size, _) => Measure(value, size),
                wrap: true,
                paragraphIndent: hanging);

            Assert.Equal(new[] { "Alpha", "Beta", "Gamma" }, plain.Lines.Select(line => line.Text).ToArray());
            Assert.Equal(0D, plain.Lines[0].OffsetX);
            Assert.Equal(2D, plain.Lines[1].OffsetX);
            Assert.Equal(2D, plain.Lines[2].OffsetX);
            Assert.Equal(7D, plain.Width);
            Assert.Equal(0D, rich.Lines[0].OffsetX);
            Assert.Equal(2D, rich.Lines[1].OffsetX);
            Assert.Equal(2D, rich.Lines[2].OffsetX);
            Assert.Equal(7D, rich.Width);
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
        public void OfficeTextLayoutEngine_KeepsFittingRemainderTogetherAfterPreferredBreak() {
            static double Measure(string? value, double size) => (value?.Length ?? 0) * size;

            IReadOnlyList<OfficeTextLine> lines = OfficeTextLayoutEngine.WrapLines("prefix-foo-bar", 1D, 7D, Measure);

            Assert.Equal(new[] { "prefix-", "foo-bar" }, lines.Select(line => line.Text).ToArray());
        }

        [Fact]
        public void OfficeTextLayoutEngine_BreaksLongWordsAtTextElementBoundaries() {
            static double Measure(string? value, double size) => string.IsNullOrEmpty(value) ? 0D : value!.Length * size;
            string eAcute = "e\u0301";
            string smile = char.ConvertFromUtf32(0x1F600);

            IReadOnlyList<OfficeTextLine> lines = OfficeTextLayoutEngine.WrapLines("A" + eAcute + smile + "B", 1D, 2D, Measure);

            Assert.Equal(new[] { "A", eAcute, smile, "B" }, lines.Select(line => line.Text).ToArray());
            Assert.DoesNotContain(lines, line => line.Text == "\u0301");
            Assert.DoesNotContain(lines, line => line.Text.Length == 1 && char.IsSurrogate(line.Text[0]));
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
        public void OfficeTextLayoutEngine_ReportsConfiguredTextLimitAsClipping() {
            static double Measure(string? value, double size) => (value?.Length ?? 0) * size;
            string oversized = new string('A', 100_001);

            OfficeTextLine endTrimmed = OfficeTextLayoutEngine.TrimLineToWidth(oversized, 1D, 200_000D, Measure, out bool endClipped);
            OfficeTextLine startTrimmed = OfficeTextLayoutEngine.TrimLineStartToWidth(oversized, 1D, 200_000D, Measure, out bool startClipped);

            Assert.True(endClipped);
            Assert.True(startClipped);
            Assert.Equal(100_000, endTrimmed.Text.Length);
            Assert.Equal(100_000, startTrimmed.Text.Length);
        }

        [Fact]
        public void OfficeTextLayoutEngine_TrimsSingleLineAtTextElementBoundaries() {
            static double Measure(string? value, double size) => string.IsNullOrEmpty(value) ? 0D : value!.Length * size;
            string eAcute = "e\u0301";
            string smile = char.ConvertFromUtf32(0x1F600);

            OfficeTextLine endTrimmed = OfficeTextLayoutEngine.TrimLineToWidth("A" + eAcute + smile + "BC", 1D, 6D, Measure, out bool endClipped);
            OfficeTextLine startTrimmed = OfficeTextLayoutEngine.TrimLineStartToWidth("XA" + eAcute + smile + "B", 1D, 6D, Measure, out bool startClipped);

            Assert.True(endClipped);
            Assert.True(startClipped);
            Assert.Equal("A" + eAcute + "...", endTrimmed.Text);
            Assert.Equal("..." + smile + "B", startTrimmed.Text);
            Assert.Contains(eAcute, endTrimmed.Text, StringComparison.Ordinal);
            Assert.Contains(smile, startTrimmed.Text, StringComparison.Ordinal);
        }

        [Fact]
        public void OfficeTextLayoutEngine_CanKeepOverflowingTextForCallerClipping() {
            double Measure(string? value, double size) => (value?.Length ?? 0) * size;

            OfficeTextBlockLayout ellipsis = OfficeTextLayoutEngine.LayoutTextBlock(
                "ABCDEFG",
                1D,
                6D,
                10D,
                lineHeightFactor: 1.2D,
                minimumFontSize: 1D,
                Measure,
                wrap: false);
            OfficeTextBlockLayout clipped = OfficeTextLayoutEngine.LayoutTextBlock(
                "ABCDEFG",
                1D,
                6D,
                10D,
                lineHeightFactor: 1.2D,
                minimumFontSize: 1D,
                Measure,
                wrap: false,
                forceSingleLine: false,
                shrinkToFit: false,
                overflowBehavior: OfficeTextOverflowBehavior.Clip);

            Assert.True(ellipsis.Clipped);
            Assert.Equal("ABC...", ellipsis.Lines[0].Text);
            Assert.Equal(6D, ellipsis.Width);
            Assert.True(clipped.Clipped);
            Assert.Equal("ABCDEFG", clipped.Lines[0].Text);
            Assert.Equal(7D, clipped.Width);
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
        public void OfficeDrawingSceneText_ShrinksWrappedTextThroughSharedRenderer() {
            var drawing = new OfficeDrawing(90D, 36D)
                .AddText(
                    "Alpha beta gamma delta epsilon",
                    0D,
                    0D,
                    90D,
                    36D,
                    new OfficeFontInfo("Aptos", 18D),
                    OfficeColor.Black,
                    wrapText: true,
                    shrinkToFit: true);

            string svg = OfficeDrawingSvgExporter.ToSvg(drawing);
            double fontSize = ExtractFirstSvgFontSize(svg);
            OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing);

            Assert.True(fontSize < 18D, "Expected scene text SVG output to shrink the authored font size.");
            Assert.True(CountPixelsNear(image, OfficeColor.Black) > 0, "Expected shrunken scene text to render in PNG output.");
        }

        [Fact]
        public void OfficeDrawingSceneText_RendersStackedTextThroughSharedRenderer() {
            var drawing = new OfficeDrawing(36D, 90D)
                .AddText(
                    "Stacked",
                    0D,
                    0D,
                    36D,
                    90D,
                    new OfficeFontInfo("Aptos", 12D),
                    OfficeColor.Black,
                    OfficeTextAlignment.Center,
                    verticalAlignment: OfficeTextVerticalAlignment.Top,
                    shrinkToFit: true,
                    stackedText: true);

            string svg = OfficeDrawingSvgExporter.ToSvg(drawing);
            OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing);

            Assert.True(CountOccurrences(svg, "<text") >= 7);
            Assert.True(CountPixelsNear(image, OfficeColor.Black) > 0, "Expected stacked scene text to render in PNG output.");
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
        public void OfficeTextLayoutEngine_ResolvesRotatedTextWidthLimitInsideBounds() {
            double unrotated = OfficeTextLayoutEngine.ResolveRotatedTextWidthLimit(80D, 40D, 12D, 0D);
            double diagonal = OfficeTextLayoutEngine.ResolveRotatedTextWidthLimit(80D, 40D, 12D, 45D);
            double vertical = OfficeTextLayoutEngine.ResolveRotatedTextWidthLimit(80D, 40D, 12D, 90D);
            double tiny = OfficeTextLayoutEngine.ResolveRotatedTextWidthLimit(6D, 6D, 20D, 45D);

            Assert.Equal(80D, unrotated);
            Assert.InRange(diagonal, 44D, 45D);
            Assert.Equal(40D, vertical, precision: 10);
            Assert.Equal(1D, tiny);
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
        public void OfficeTextLayoutEngine_CanKeepOverflowingRichTextForCallerClipping() {
            double Measure(string? value, double size) => (value?.Length ?? 0) * size;
            OfficeRichTextRun strong = new OfficeRichTextRun("Overflowing", 1D, OfficeColor.Red, bold: true);
            OfficeRichTextRun accent = new OfficeRichTextRun(" rich", 1D, OfficeColor.Blue, italic: true);

            OfficeRichTextBlockLayout ellipsis = OfficeTextLayoutEngine.LayoutRichTextBlock(
                new[] { strong, accent },
                10D,
                10D,
                lineHeightFactor: 1.2D,
                Measure,
                wrap: false);
            OfficeRichTextBlockLayout clipped = OfficeTextLayoutEngine.LayoutRichTextBlock(
                new[] { strong, accent },
                10D,
                10D,
                lineHeightFactor: 1.2D,
                Measure,
                wrap: false,
                shrinkToFit: false,
                minimumFontSize: 1D,
                overflowBehavior: OfficeTextOverflowBehavior.Clip);

            Assert.True(ellipsis.Clipped);
            Assert.Equal("Overflo...", string.Concat(ellipsis.Lines[0].Segments.Select(segment => segment.Text)));
            Assert.True(ellipsis.Width <= 10.01D);
            Assert.True(clipped.Clipped);
            Assert.Equal("Overflowing rich", string.Concat(clipped.Lines[0].Segments.Select(segment => segment.Text)));
            Assert.True(clipped.Width > 10D);
            Assert.True(clipped.Lines[0].Segments[0].Bold);
            Assert.True(clipped.Lines[0].Segments[1].Italic);
        }

        [Fact]
        public void OfficeTextLayoutEngine_LayoutsRichTextAtTextElementBoundaries() {
            double Measure(string? value, double size) => (value?.Length ?? 0) * size;
            string eAcute = "e\u0301";
            string smile = char.ConvertFromUtf32(0x1F600);

            OfficeRichTextBlockLayout wrapped = OfficeTextLayoutEngine.LayoutRichTextBlock(
                new[] { new OfficeRichTextRun("A" + smile + "B", 1D, OfficeColor.Red, bold: true) },
                2D,
                10D,
                lineHeightFactor: 1D,
                Measure,
                wrap: true,
                shrinkToFit: false,
                minimumFontSize: 1D,
                overflowBehavior: OfficeTextOverflowBehavior.Clip);
            OfficeRichTextBlockLayout ellipsis = OfficeTextLayoutEngine.LayoutRichTextBlock(
                new[] { new OfficeRichTextRun("A" + eAcute + smile + "BC", 1D, OfficeColor.Blue, italic: true) },
                6D,
                10D,
                lineHeightFactor: 1D,
                Measure,
                wrap: false);

            Assert.Equal(new[] { "A", smile, "B" }, wrapped.Lines.Select(line => string.Concat(line.Segments.Select(segment => segment.Text))).ToArray());
            Assert.True(wrapped.Lines.All(line => line.Segments.All(segment => segment.Bold)));
            Assert.True(ellipsis.Clipped);
            Assert.Equal("A" + eAcute + "...", string.Concat(ellipsis.Lines[0].Segments.Select(segment => segment.Text)));
            Assert.Contains(eAcute, ellipsis.Lines[0].Segments[0].Text, StringComparison.Ordinal);
            Assert.DoesNotContain(smile, ellipsis.Lines[0].Segments[0].Text, StringComparison.Ordinal);
            Assert.True(ellipsis.Lines[0].Segments.Last().Italic);
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
        public void OfficeTextLayoutEngine_UsesPerLineHeightsForMixedSizeRichText() {
            double Measure(string? value, double size) => (value?.Length ?? 0) * size;

            OfficeRichTextBlockLayout layout = OfficeTextLayoutEngine.LayoutRichTextBlock(
                new[] {
                    new OfficeRichTextRun("Big ", 20D, OfficeColor.Black),
                    new OfficeRichTextRun("small", 8D, OfficeColor.Blue)
                },
                75D,
                34D,
                lineHeightFactor: 1.2D,
                Measure,
                wrap: true,
                shrinkToFit: false,
                minimumFontSize: 1D,
                overflowBehavior: OfficeTextOverflowBehavior.Clip);

            Assert.False(layout.Clipped);
            Assert.Equal(2, layout.Lines.Count);
            Assert.Equal(24D, layout.LineHeight);
            Assert.Equal(24D, layout.Lines[0].LineHeight);
            Assert.Equal(10D, layout.Lines[1].LineHeight);
            Assert.Equal(34D, layout.Height);
            Assert.Equal(new[] { "Big", "small" }, layout.Lines.Select(line => string.Concat(line.Segments.Select(segment => segment.Text))).ToArray());
        }

        [Fact]
        public void OfficeTextLayoutEngine_RichTextHonorsCancellation() {
            using var cancellation = new System.Threading.CancellationTokenSource();
            cancellation.Cancel();

            Assert.Throws<OperationCanceledException>(() =>
                OfficeTextLayoutEngine.LayoutRichTextBlock(
                    new[] {
                        new OfficeRichTextRun("Large rich text", 12D, OfficeColor.Black)
                    },
                    100D,
                    100D,
                    lineHeightFactor: 1.2D,
                    (value, size, _) => (value?.Length ?? 0) * size,
                    wrap: true,
                    shrinkToFit: false,
                    minimumFontSize: 1D,
                    overflowBehavior: OfficeTextOverflowBehavior.Clip,
                    paragraphIndent: null,
                    cancellationToken: cancellation.Token));
        }

        [Fact]
        public void OfficeTextLayoutEngine_CapsRichTextLinesBeforeHeightClipping() {
            string text = string.Join("\n", Enumerable.Repeat("A", 5_000));

            OfficeRichTextBlockLayout layout = OfficeTextLayoutEngine.LayoutRichTextBlock(
                new[] { new OfficeRichTextRun(text, 1D, OfficeColor.Black) },
                10D,
                100_000D,
                lineHeightFactor: 1D,
                static (value, size, _) => (value?.Length ?? 0) * size,
                wrap: true,
                shrinkToFit: false,
                minimumFontSize: 1D,
                overflowBehavior: OfficeTextOverflowBehavior.Clip);

            Assert.True(layout.Clipped);
            Assert.Equal(4_096, layout.Lines.Count);
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
        public void OfficeRasterCanvas_LongDashedLinesReachTheVisibleEndpoint() {
            OfficeRasterImage dashed = new OfficeRasterImage(4000, 8, OfficeColor.Transparent);
            OfficeRasterImage patterned = new OfficeRasterImage(4000, 8, OfficeColor.Transparent);

            new OfficeRasterCanvas(dashed).DrawDashedLine(
                0D, 3D, 3999D, 3D, OfficeColor.Black, 2D, dashLength: 23.04D, gapLength: 11.52D);
            new OfficeRasterCanvas(patterned).DrawPatternedLine(
                0D, 4D, 3999D, 4D, OfficeColor.Black, 2D, new[] { 23.04D, 11.52D });

            Assert.Contains(Enumerable.Range(3900, 100), x => dashed.GetPixel(x, 3).A > 0);
            Assert.Contains(Enumerable.Range(3900, 100), x => patterned.GetPixel(x, 4).A > 0);
        }

        [Fact]
        public async System.Threading.Tasks.Task OfficeRasterCanvas_ZeroGapDashedEllipseAdvancesPastNearBoundaryPhase() {
            OfficeRasterImage image = new OfficeRasterImage(32, 32, OfficeColor.Transparent);
            double firstSegmentLength = Math.Sqrt(200D);

            System.Threading.Tasks.Task render = System.Threading.Tasks.Task.Run(() => new OfficeRasterCanvas(image).DrawDashedEllipse(
                16D,
                16D,
                10D,
                10D,
                OfficeColor.Black,
                thickness: 1D,
                dashLength: firstSegmentLength + 0.0000000005D,
                gapLength: 0D,
                segments: 4));

            System.Threading.Tasks.Task completed = await System.Threading.Tasks.Task.WhenAny(
                render,
                System.Threading.Tasks.Task.Delay(TimeSpan.FromSeconds(2)));
            Assert.Same(render, completed);
            await render;
            Assert.True(CountPaintedPixels(image) > 0);
        }

        [Fact]
        public void OfficeRasterCanvas_ClipsHugeDashedAndPatternedLinesBeforeIteration() {
            OfficeRasterImage dashed = new OfficeRasterImage(32, 8, OfficeColor.Transparent);
            OfficeRasterImage patterned = new OfficeRasterImage(32, 8, OfficeColor.Transparent);

            new OfficeRasterCanvas(dashed).DrawDashedLine(-1_000_000_000D, 3D, 1_000_000_000D, 3D, OfficeColor.Black, 1D, 2D, 2D);
            new OfficeRasterCanvas(patterned).DrawPatternedLine(-1_000_000_000D, 4D, 1_000_000_000D, 4D, OfficeColor.Black, 1D, new[] { 2D, 2D });

            Assert.Contains(Enumerable.Range(0, dashed.Width), x => dashed.GetPixel(x, 3).A > 0);
            Assert.Contains(Enumerable.Range(0, patterned.Width), x => patterned.GetPixel(x, 4).A > 0);
        }

        [Fact]
        public void OfficeRasterCanvas_Quantizes_Tiny_Dash_Patterns_To_Raster_Resolution() {
            OfficeRasterImage dashed = new OfficeRasterImage(32, 8, OfficeColor.Transparent);
            OfficeRasterImage patterned = new OfficeRasterImage(32, 8, OfficeColor.Transparent);

            new OfficeRasterCanvas(dashed).DrawDashedLine(0D, 3D, 31D, 3D, OfficeColor.Black, 1D, 0.00000001D, 0.00000001D);
            new OfficeRasterCanvas(patterned).DrawPatternedLine(0D, 4D, 31D, 4D, OfficeColor.Black, 1D, new[] { 0.00000001D, 0.00000001D });

            Assert.Contains(Enumerable.Range(0, dashed.Width), x => dashed.GetPixel(x, 3).A > 0);
            Assert.Contains(Enumerable.Range(0, patterned.Width), x => patterned.GetPixel(x, 4).A > 0);
        }

        [Fact]
        public void OfficeRasterCanvas_TinyDashNormalizationDoesNotOverflowLargeGaps() {
            OfficeRasterImage dashed = new OfficeRasterImage(32, 8, OfficeColor.Transparent);
            OfficeRasterImage patterned = new OfficeRasterImage(32, 8, OfficeColor.Transparent);

            new OfficeRasterCanvas(dashed).DrawDashedLine(
                0D, 3D, 31D, 3D, OfficeColor.Black, 1D, dashLength: 1e-300D, gapLength: 1e100D);
            new OfficeRasterCanvas(patterned).DrawPatternedLine(
                0D, 4D, 31D, 4D, OfficeColor.Black, 1D, new[] { 1e-300D, 1e100D });

            Assert.Contains(Enumerable.Range(0, dashed.Width), x => dashed.GetPixel(x, 3).A > 0);
            Assert.Contains(Enumerable.Range(0, patterned.Width), x => patterned.GetPixel(x, 4).A > 0);
        }

        [Fact]
        public void OfficeRasterCanvas_FiniteDashLengthsRenderWhenTheirCycleOverflows() {
            OfficeRasterImage dashed = new OfficeRasterImage(32, 8, OfficeColor.Transparent);
            OfficeRasterImage patterned = new OfficeRasterImage(32, 8, OfficeColor.Transparent);

            new OfficeRasterCanvas(dashed).DrawDashedLine(
                0D, 3D, 31D, 3D, OfficeColor.Black, 1D, double.MaxValue, double.MaxValue);
            new OfficeRasterCanvas(patterned).DrawPatternedLine(
                0D, 4D, 31D, 4D, OfficeColor.Black, 1D, new[] { double.MaxValue, double.MaxValue });

            Assert.Contains(Enumerable.Range(0, dashed.Width), x => dashed.GetPixel(x, 3).A > 0);
            Assert.Contains(Enumerable.Range(0, patterned.Width), x => patterned.GetPixel(x, 4).A > 0);
        }

        [Fact]
        public void OfficeRasterCanvas_RejectsNonFiniteDashedLineLengths() {
            OfficeRasterImage image = new OfficeRasterImage(32, 8, OfficeColor.Transparent);
            OfficeRasterCanvas canvas = new OfficeRasterCanvas(image);
            var points = new[] { new OfficePoint(0D, 3D), new OfficePoint(31D, 3D) };

            canvas.DrawDashedLine(0D, 3D, 31D, 3D, OfficeColor.Black, dashLength: double.PositiveInfinity);
            canvas.DrawDashedLine(0D, 3D, 31D, 3D, OfficeColor.Black, gapLength: double.NaN);
            canvas.DrawDashedPolyline(points, OfficeColor.Black, dashLength: double.PositiveInfinity);
            canvas.DrawDashedPolyline(points, OfficeColor.Black, gapLength: double.PositiveInfinity);

            Assert.All(image.GetPixels(), value => Assert.Equal(0, value));
        }

        [Fact]
        public void OfficeRasterCanvas_OddDashPatternsPreserveAlternatingParityAcrossWrap() {
            OfficeRasterImage odd = new OfficeRasterImage(32, 8, OfficeColor.Transparent);
            OfficeRasterImage duplicated = new OfficeRasterImage(32, 8, OfficeColor.Transparent);

            new OfficeRasterCanvas(odd).DrawPatternedLine(
                0D, 3D, 31D, 3D, OfficeColor.Black, 1D, new[] { 2D, 1D, 3D });
            new OfficeRasterCanvas(duplicated).DrawPatternedLine(
                0D, 3D, 31D, 3D, OfficeColor.Black, 1D, new[] { 2D, 1D, 3D, 2D, 1D, 3D });

            Assert.Equal(duplicated.GetPixels(), odd.GetPixels());
        }

        [Fact]
        public void OfficeRasterCanvas_ClippedTinyDashPhaseMatchesQuantizedPattern() {
            OfficeRasterImage tinyDashed = new OfficeRasterImage(32, 8, OfficeColor.Transparent);
            OfficeRasterImage quantizedDashed = new OfficeRasterImage(32, 8, OfficeColor.Transparent);
            OfficeRasterImage tinyPatterned = new OfficeRasterImage(32, 8, OfficeColor.Transparent);
            OfficeRasterImage quantizedPatterned = new OfficeRasterImage(32, 8, OfficeColor.Transparent);

            new OfficeRasterCanvas(tinyDashed).DrawDashedLine(-3.625D, 3D, 31D, 3D, OfficeColor.Black, 1D, 0.00000001D, 0.00000001D);
            new OfficeRasterCanvas(quantizedDashed).DrawDashedLine(-3.625D, 3D, 31D, 3D, OfficeColor.Black, 1D, 0.25D, 0.25D);
            new OfficeRasterCanvas(tinyPatterned).DrawPatternedLine(-3.625D, 4D, 31D, 4D, OfficeColor.Black, 1D, new[] { 0.00000001D, 0.00000001D });
            new OfficeRasterCanvas(quantizedPatterned).DrawPatternedLine(-3.625D, 4D, 31D, 4D, OfficeColor.Black, 1D, new[] { 0.25D, 0.25D });

            Assert.Equal(quantizedDashed.GetPixels(), tinyDashed.GetPixels());
            Assert.Equal(quantizedPatterned.GetPixels(), tinyPatterned.GetPixels());
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
        public void OfficeDrawingRasterRenderer_KeepsZeroWidthStrokesInvisible() {
            OfficeDrawing drawing = new OfficeDrawing(32, 24);
            OfficeShape shape = OfficeShape.Rectangle(20, 12);
            shape.StrokeColor = OfficeColor.Red;
            shape.StrokeWidth = 0D;
            drawing.AddShape(shape, 6, 6);

            OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing);

            Assert.Equal(0, CountPaintedPixels(image));
        }

        [Fact]
        public void OfficeDrawingRasterRenderer_AppliesFillOpacityToGradients() {
            OfficeDrawing drawing = new OfficeDrawing(32, 24);
            OfficeShape shape = OfficeShape.Rectangle(24, 16);
            shape.FillGradient = new OfficeLinearGradient(
                0,
                0.5,
                1,
                0.5,
                new[] {
                    new OfficeGradientStop(0, OfficeColor.Red),
                    new OfficeGradientStop(0.5, OfficeColor.Lime),
                    new OfficeGradientStop(1, OfficeColor.Blue)
                });
            shape.FillOpacity = 0.5D;
            drawing.AddShape(shape, 4, 4);

            OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing);
            OfficeColor middle = image.GetPixel(16, 12);

            Assert.InRange(middle.A, 100, 155);
            Assert.True(middle.G > middle.R);
            Assert.True(middle.G > middle.B);
        }

        [Fact]
        public void OfficeDrawingRasterRenderer_FillsRoundedRectangleGradientsInsideRoundedContour() {
            OfficeDrawing drawing = new OfficeDrawing(48, 36);
            OfficeShape shape = OfficeShape.RoundedRectangle(28, 18, 7);
            shape.FillGradient = OfficeLinearGradient.Horizontal(OfficeColor.Red, OfficeColor.Blue);
            drawing.AddShape(shape, 4, 4);

            OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing);

            Assert.Equal(0, image.GetPixel(4, 4).A);
            OfficeColor middle = image.GetPixel(18, 13);
            Assert.True(middle.A > 200);
            Assert.True(middle.R > 20 || middle.B > 20);
        }

        [Fact]
        public void OfficeDrawingRasterRenderer_FillsEllipseGradientsInsideEllipseContour() {
            OfficeDrawing drawing = new OfficeDrawing(48, 36);
            OfficeShape shape = OfficeShape.Ellipse(28, 18);
            shape.FillGradient = OfficeLinearGradient.Horizontal(OfficeColor.Red, OfficeColor.Blue);
            drawing.AddShape(shape, 4, 4);

            OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing);

            Assert.Equal(0, image.GetPixel(4, 4).A);
            OfficeColor middle = image.GetPixel(18, 13);
            Assert.True(middle.A > 200);
            Assert.True(middle.R > 20 || middle.B > 20);
        }

        [Fact]
        public void OfficeDrawingRasterRenderer_FillsPolygonGradients() {
            OfficeDrawing drawing = new OfficeDrawing(48, 28);
            OfficeShape shape = OfficeShape.Polygon(
                new OfficePoint(0, 0),
                new OfficePoint(32, 0),
                new OfficePoint(32, 18),
                new OfficePoint(0, 18));
            shape.FillGradient = OfficeLinearGradient.Horizontal(OfficeColor.Red, OfficeColor.Blue);
            drawing.AddShape(shape, 4, 4);

            OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing);

            OfficeColor left = image.GetPixel(8, 12);
            OfficeColor right = image.GetPixel(32, 12);
            Assert.True(left.R > left.B, $"Expected polygon gradient left side to keep the red stop, got {left}.");
            Assert.True(right.B > right.R, $"Expected polygon gradient right side to keep the blue stop, got {right}.");
        }

        [Fact]
        public void OfficeDrawingRasterRenderer_FillsPathGradients() {
            OfficeDrawing drawing = new OfficeDrawing(48, 28);
            OfficeShape shape = OfficeShape.Path(
                OfficePathCommand.MoveTo(0, 0),
                OfficePathCommand.LineTo(32, 0),
                OfficePathCommand.LineTo(32, 18),
                OfficePathCommand.LineTo(0, 18),
                OfficePathCommand.Close());
            shape.FillGradient = OfficeLinearGradient.Horizontal(OfficeColor.Red, OfficeColor.Blue);
            drawing.AddShape(shape, 4, 4);

            OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing);

            OfficeColor left = image.GetPixel(8, 12);
            OfficeColor right = image.GetPixel(32, 12);
            Assert.True(left.R > left.B, $"Expected path gradient left side to keep the red stop, got {left}.");
            Assert.True(right.B > right.R, $"Expected path gradient right side to keep the blue stop, got {right}.");
        }

        [Fact]
        public void OfficeDrawingRasterRenderer_PreservesGradientOnTransformedShapes() {
            OfficeDrawing drawing = new OfficeDrawing(64, 48);
            OfficeShape shape = OfficeShape.Rectangle(34, 20);
            shape.FillGradient = OfficeLinearGradient.Horizontal(OfficeColor.Red, OfficeColor.Blue);
            shape.Transform = OfficeTransform.RotateDegrees(20D, 17D, 10D);
            drawing.AddShape(shape, 14, 12);

            OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing);

            OfficeColor left = image.GetPixel(20, 25);
            OfficeColor right = image.GetPixel(45, 23);
            Assert.True(left.R > left.B, $"Expected transformed gradient left side to keep the red stop, got {left}.");
            Assert.True(right.B > right.R, $"Expected transformed gradient right side to keep the blue stop, got {right}.");
        }

        [Fact]
        public void OfficeDrawingRasterRenderer_RotatesLocalGradientWithShapeTransform() {
            OfficeDrawing drawing = new OfficeDrawing(100, 80);
            OfficeShape shape = OfficeShape.Rectangle(40, 20);
            shape.FillGradient = OfficeLinearGradient.Horizontal(OfficeColor.Red,
                OfficeColor.Blue);
            shape.Transform = OfficeTransform.RotateDegrees(90D, 20D, 10D);
            drawing.AddShape(shape, 30, 20);

            OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing);

            OfficeColor top = image.GetPixel(50, 12);
            OfficeColor bottom = image.GetPixel(50, 48);
            Assert.True(top.R > top.B,
                $"Expected the transformed local gradient to start at the rotated top edge, got {top}.");
            Assert.True(bottom.B > bottom.R,
                $"Expected the transformed local gradient to end at the rotated bottom edge, got {bottom}.");
        }

        [Fact]
        public void OfficeLinearGradient_FromTransformedAnglePreservesArbitraryDestinationDirection() {
            const double expectedDegrees = 37D;
            const double width = 100D;
            const double height = 40D;
            var frame = new OfficeImageFrameTransform(31D, width / 2D,
                height / 2D);
            OfficeTransform transform = frame.CreateDestinationTransform();
            OfficeLinearGradient gradient = OfficeLinearGradient.FromTransformedAngle(
                new[] {
                    new OfficeGradientStop(0D, OfficeColor.Red),
                    new OfficeGradientStop(1D, OfficeColor.Blue)
                },
                expectedDegrees,
                width,
                height,
                transform);

            OfficePoint start = transform.TransformPoint(new OfficePoint(
                gradient.StartX * width, gradient.StartY * height));
            OfficePoint end = transform.TransformPoint(new OfficePoint(
                gradient.EndX * width, gradient.EndY * height));
            double actualDegrees = Math.Atan2(end.Y - start.Y, end.X - start.X)
                * 180D / Math.PI;
            if (actualDegrees < 0D) actualDegrees += 360D;

            Assert.InRange(actualDegrees, expectedDegrees - 0.001D,
                expectedDegrees + 0.001D);
        }

        [Fact]
        public void OfficeDrawingRasterRenderer_RendersRoundedRectanglesAndShapeShadows() {
            OfficeDrawing drawing = new OfficeDrawing(48, 36);
            OfficeShape shape = OfficeShape.RoundedRectangle(24, 16, 6);
            shape.FillColor = OfficeColor.Red;
            shape.Shadow = new OfficeShadow(OfficeColor.Black, 0.5D, 8D, 6D);
            drawing.AddShape(shape, 4, 4);

            OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing);

            Assert.Equal(0, image.GetPixel(4, 4).A);
            Assert.True(image.GetPixel(16, 12).A > 200);
            Assert.True(image.GetPixel(34, 24).A > 0, "Expected the shape shadow to render behind the rounded rectangle.");
        }

        [Fact]
        public void OfficeDrawingRasterRenderer_AppliesRasterClipPaths() {
            OfficeDrawing drawing = new OfficeDrawing(32, 24);
            OfficeShape shape = OfficeShape.Rectangle(24, 16);
            shape.FillColor = OfficeColor.Red;
            shape.ClipPath = OfficeClipPath.RoundedRectangle(24, 16, 7);
            drawing.AddShape(shape, 4, 4);

            OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing);

            Assert.Equal(0, image.GetPixel(4, 4).A);
            Assert.True(image.GetPixel(16, 12).A > 200);
        }

        [Fact]
        public void OfficeDrawingRasterRenderer_HonorsNonRectangularGroupClipPaths() {
            OfficeDrawing child = new OfficeDrawing(24, 16);
            OfficeShape fill = OfficeShape.Rectangle(24, 16);
            fill.FillColor = OfficeColor.Red;
            child.AddShape(fill, 0, 0);
            OfficeDrawing drawing = new OfficeDrawing(32, 24);
            drawing.AddClippedDrawing(child, 4, 4, OfficeClipPath.RoundedRectangle(24, 16, 7));

            OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing);

            Assert.Equal(0, image.GetPixel(4, 4).A);
            Assert.True(image.GetPixel(16, 12).A > 200);
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
        public void OfficeRasterCanvas_SkipsPatternedPolylineWithNonFiniteSegmentLength() {
            OfficeRasterImage image = new OfficeRasterImage(8, 8, OfficeColor.Transparent);
            new OfficeRasterCanvas(image).DrawPatternedPolyline(
                new[] {
                    new OfficePoint(1, 1),
                    new OfficePoint(double.PositiveInfinity, 1)
                },
                OfficeColor.Black,
                1D,
                new[] { 1D, 1D });

            Assert.Equal(0, CountPaintedPixels(image));
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
        public void OfficeRasterCanvas_ScalesTransparentImageEdgesWithoutDarkeningColor() {
            OfficeRasterImage source = new OfficeRasterImage(2, 2, OfficeColor.Transparent);
            source.SetPixel(0, 0, OfficeColor.FromRgb(0, 255, 0));
            OfficeRasterImage target = new OfficeRasterImage(8, 8, OfficeColor.Transparent);
            OfficeRasterCanvas canvas = new OfficeRasterCanvas(target);

            canvas.DrawImage(source, 0, 0, 8, 8);

            OfficeColor edge = target.GetPixel(4, 0);
            Assert.InRange(edge.A, 1, 254);
            Assert.True(edge.G > 220, "Expected transparent-edge image scaling to preserve premultiplied source color instead of blending toward transparent black.");
            Assert.True(edge.R < 10);
            Assert.True(edge.B < 10);
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
        public void OfficeRasterCanvas_DrawTextTrimsAtTextElementBoundaries() {
            if (OfficeTrueTypeFont.TryLoadDefault() == null) {
                return;
            }

            string eAcute = "e\u0301";
            string smile = char.ConvertFromUtf32(0x1F600);
            const double fontSize = 20D;
            double availableWidth = Math.Ceiling(new OfficeRasterCanvas(new OfficeRasterImage(1, 1, OfficeColor.Transparent)).MeasureText("A" + eAcute + "...", fontSize)) + 0.5D;
            double boxWidth = availableWidth + 6D;
            OfficeRasterImage clipped = new OfficeRasterImage(120, 40, OfficeColor.Transparent);
            OfficeRasterImage expected = new OfficeRasterImage(120, 40, OfficeColor.Transparent);

            new OfficeRasterCanvas(clipped).DrawText("A" + eAcute + smile + "BC", 0D, 0D, boxWidth, 32D, OfficeColor.Black, fontSize);
            new OfficeRasterCanvas(expected).DrawText("A" + eAcute + "...", 0D, 0D, boxWidth, 32D, OfficeColor.Black, fontSize);

            AssertRasterImagesEqual(expected, clipped);
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
        public void OfficeTextBlockRenderer_DrawsRasterTextBoxWithBackground() {
            OfficeRasterImage image = new OfficeRasterImage(140, 80, OfficeColor.Transparent);
            OfficeRasterCanvas canvas = new OfficeRasterCanvas(image);
            OfficeTextBlockRenderPlan plan = OfficeTextBlockRenderPlan.CreateFittedFromCenter(
                "Shared text",
                12D,
                centerX: 70D,
                centerY: 40D,
                width: 90D,
                height: 36D,
                canvas.MeasureText,
                OfficeTextAlignment.Center,
                OfficeTextVerticalAlignment.Center,
                lineHeightFactor: 1.2D,
                minimumFontSize: 8D);

            OfficeTextBlockRenderer.DrawRasterTextBox(
                canvas,
                plan,
                OfficeColor.Black,
                bold: true,
                backgroundColor: OfficeColor.FromRgb(255, 230, 128),
                backgroundPaddingX: 4D,
                backgroundPaddingY: 3D);

            Assert.True(CountPixelsNear(image, OfficeColor.FromRgb(255, 230, 128)) > 100, "Expected the shared text-box helper to paint the background.");
            Assert.True(CountPixelsNear(image, OfficeColor.Black) > 40, "Expected the shared text-box helper to paint text.");
        }

        [Fact]
        public void OfficeTextBlockRenderer_DrawsRasterRichTextBlockWithRunStyles() {
            OfficeRasterImage image = new OfficeRasterImage(140, 72, OfficeColor.Transparent);
            OfficeRasterCanvas canvas = new OfficeRasterCanvas(image);
            var line = new OfficeRichTextLine(new[] {
                new OfficeRichTextSegment("Red", canvas.MeasureText("Red", 14D), 14D, OfficeColor.Red, bold: true, italic: false, underline: true, fontFamily: "Aptos", backgroundColor: OfficeColor.Yellow),
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
            Assert.True(CountPixelsNear(image, OfficeColor.Yellow) > 20, "Expected the first rich text run background to render yellow.");
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
        public void OfficeTextBlockRenderer_AppendsSvgTextBlockWithLineOffsets() {
            var builder = new System.Text.StringBuilder();
            var layout = new OfficeTextBlockLayout(
                new[] {
                    new OfficeTextLine("First", 20D),
                    new OfficeTextLine("Rest", 20D, 10D)
                },
                fontSize: 10D,
                lineHeight: 12D,
                width: 30D,
                height: 24D);

            builder.AppendSvgTextBlock(
                layout,
                left: 0D,
                top: 0D,
                width: 100D,
                height: 40D,
                color: OfficeColor.Black,
                fontFamily: "Aptos");

            string svg = builder.ToString();
            Assert.Contains("<text x=\"0\"", svg, StringComparison.Ordinal);
            Assert.Contains("<text x=\"10\"", svg, StringComparison.Ordinal);
            Assert.Contains(">First</text>", svg, StringComparison.Ordinal);
            Assert.Contains(">Rest</text>", svg, StringComparison.Ordinal);
        }

        [Fact]
        public void OfficeTextBlockRenderer_AppendsSvgJustifiedTextBlockWithTextLength() {
            var builder = new System.Text.StringBuilder();
            var layout = new OfficeTextBlockLayout(
                new[] {
                    new OfficeTextLine("Alpha Beta", 48D),
                    new OfficeTextLine("Gamma", 25D)
                },
                fontSize: 10D,
                lineHeight: 12D,
                width: 48D,
                height: 24D);

            builder.AppendSvgTextBlock(
                layout,
                left: 0D,
                top: 0D,
                width: 100D,
                height: 40D,
                color: OfficeColor.Black,
                fontFamily: "Aptos",
                horizontalAlignment: OfficeTextAlignment.Justify);

            string svg = builder.ToString();
            Assert.Contains("<text x=\"0\"", svg, StringComparison.Ordinal);
            Assert.Contains("textLength=\"100\" lengthAdjust=\"spacing\"", svg, StringComparison.Ordinal);
            Assert.Contains(">Alpha Beta</text>", svg, StringComparison.Ordinal);
            Assert.Contains(">Gamma</text>", svg, StringComparison.Ordinal);
            Assert.Equal(1, svg.Split(new[] { "textLength=" }, StringSplitOptions.None).Length - 1);
        }

        [Fact]
        public void OfficeTextBlockRenderer_AppendsSvgJustifiedRichTextBlockWithTextLengthAndTspans() {
            var builder = new System.Text.StringBuilder();
            var firstLine = new OfficeRichTextLine(new[] {
                new OfficeRichTextSegment("Alpha ", width: 30D, fontSize: 10D, color: OfficeColor.Red, bold: true, italic: false, underline: false, fontFamily: "Aptos"),
                new OfficeRichTextSegment("Beta", width: 20D, fontSize: 12D, color: OfficeColor.Blue, bold: false, italic: true, underline: true, fontFamily: "Aptos", strikethrough: true)
            });
            var secondLine = new OfficeRichTextLine(new[] {
                new OfficeRichTextSegment("Tail", width: 18D, fontSize: 10D, color: OfficeColor.Black, bold: false, italic: false, underline: false, fontFamily: "Aptos")
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
                height: 40D,
                horizontalAlignment: OfficeTextAlignment.Justify);

            string svg = builder.ToString();
            Assert.Contains("<text x=\"0\" y=\"11.08\"", svg, StringComparison.Ordinal);
            Assert.Contains("textLength=\"100\" lengthAdjust=\"spacing\"", svg, StringComparison.Ordinal);
            Assert.Contains("<tspan", svg, StringComparison.Ordinal);
            Assert.Contains("font-weight=\"700\"", svg, StringComparison.Ordinal);
            Assert.Contains("font-style=\"italic\"", svg, StringComparison.Ordinal);
            Assert.Contains("text-decoration=\"underline line-through\"", svg, StringComparison.Ordinal);
            Assert.Contains(">Alpha </tspan>", svg, StringComparison.Ordinal);
            Assert.Contains(">Beta</tspan>", svg, StringComparison.Ordinal);
            Assert.Contains(">Tail</text>", svg, StringComparison.Ordinal);
            Assert.Equal(1, svg.Split(new[] { "textLength=" }, StringSplitOptions.None).Length - 1);
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
            Assert.Equal("<text x=\"50\" y=\"12\" font-family=\"Aptos\" font-size=\"10\" text-anchor=\"middle\" fill=\"#010203\" fill-opacity=\"0.502\" xml:space=\"preserve\" font-weight=\"700\" font-style=\"italic\" text-decoration=\"underline line-through\" transform=\"rotate(15 50 20)\">A&amp;B<tspan x=\"50\" dy=\"14\">Beta</tspan></text>", svg);
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
        public void OfficeTextBlockRenderer_AppendsSvgRichTextSegmentWithPreservedBoundaryWhitespace() {
            var builder = new System.Text.StringBuilder();
            var segment = new OfficeRichTextSegment(
                " spaced",
                width: 34D,
                fontSize: 10D,
                color: OfficeColor.Blue,
                bold: false,
                italic: true,
                underline: false,
                fontFamily: "Aptos");

            builder.AppendSvgRichTextSegment(segment, 5D, 12D);

            string svg = builder.ToString();
            Assert.Contains("xml:space=\"preserve\"", svg);
            Assert.Contains("> spaced</text>", svg);
        }

        [Fact]
        public void OfficeTextBlockRenderer_AppendsSvgRichTextBlockWithSharedPlacement() {
            var builder = new System.Text.StringBuilder();
            var firstLine = new OfficeRichTextLine(new[] {
                new OfficeRichTextSegment("Red", width: 20D, fontSize: 10D, color: OfficeColor.Red, bold: true, italic: false, underline: true, fontFamily: "Aptos", backgroundColor: OfficeColor.Yellow),
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
            Assert.Contains("<rect", svg, StringComparison.Ordinal);
            Assert.Contains("fill=\"#FFFF00\"", svg, StringComparison.OrdinalIgnoreCase);
            Assert.True(svg.IndexOf("<rect", StringComparison.Ordinal) < svg.IndexOf(">Red</text>", StringComparison.Ordinal));
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

        [Fact]
        public void OfficeTextBlockRenderer_WritesSvgTextBoxWithBackgroundAndAdapterAttributes() {
            var layout = new OfficeTextBlockLayout(
                new[] {
                    new OfficeTextLine("Shared", 36D)
                },
                fontSize: 10D,
                lineHeight: 12D,
                width: 36D,
                height: 12D);
            OfficeTextBlockRenderPlan plan = OfficeTextBlockRenderPlan.CreateFromCenter(
                layout,
                centerX: 50D,
                centerY: 25D,
                width: 80D,
                height: 30D,
                OfficeTextAlignment.Center,
                OfficeTextVerticalAlignment.Center);
            var builder = new System.Text.StringBuilder();
            using (var writer = System.Xml.XmlWriter.Create(
                builder,
                new System.Xml.XmlWriterSettings {
                    ConformanceLevel = System.Xml.ConformanceLevel.Fragment,
                    OmitXmlDeclaration = true
                })) {
                OfficeTextBlockRenderer.WriteSvgTextBox(
                    writer,
                    plan,
                    OfficeColor.FromRgb(1, 2, 3),
                    "Aptos",
                    bold: true,
                    rotationDegrees: 12D,
                    rotationCenterX: 50D,
                    rotationCenterY: 25D,
                    svgNamespace: "http://www.w3.org/2000/svg",
                    backgroundColor: OfficeColor.FromRgba(255, 230, 128, 200),
                    backgroundPaddingX: 4D,
                    backgroundPaddingY: 3D,
                    configureTextAttributes: textWriter => textWriter.WriteAttributeString("data-officeimo-text", "true"),
                    configureBackgroundAttributes: backgroundWriter => backgroundWriter.WriteAttributeString("data-officeimo-background", "true"));
            }

            string svg = builder.ToString();
            Assert.Contains("<rect", svg);
            Assert.Contains("data-officeimo-background=\"true\"", svg);
            Assert.Contains("fill=\"#FFE680\"", svg);
            Assert.Contains("fill-opacity=\"0.784\"", svg);
            Assert.Contains("transform=\"rotate(12 50 25)\"", svg);
            Assert.Contains("<text", svg);
            Assert.Contains("data-officeimo-text=\"true\"", svg);
            Assert.Contains("font-family=\"Aptos\"", svg);
            Assert.Contains("font-weight=\"700\"", svg);
            Assert.Contains(">Shared</tspan>", svg);
        }

        private static int CountPaintedPixels(OfficeRasterImage image) =>
            Enumerable.Range(0, image.Width * image.Height)
                .Count(index => image.GetPixel(index % image.Width, index / image.Width).A > 0);

        private static byte[] CreateBmp24(int width, int height, IReadOnlyList<OfficeColor> pixels, bool topDown = false) {
            int rowStride = ((width * 24) + 31) / 32 * 4;
            int pixelOffset = 54;
            byte[] bytes = new byte[pixelOffset + (rowStride * height)];
            bytes[0] = (byte)'B';
            bytes[1] = (byte)'M';
            WriteInt32LittleEndian(bytes, 2, bytes.Length);
            WriteInt32LittleEndian(bytes, 10, pixelOffset);
            WriteInt32LittleEndian(bytes, 14, 40);
            WriteInt32LittleEndian(bytes, 18, width);
            WriteInt32LittleEndian(bytes, 22, topDown ? -height : height);
            WriteUInt16LittleEndian(bytes, 26, 1);
            WriteUInt16LittleEndian(bytes, 28, 24);

            for (int y = 0; y < height; y++) {
                int sourceY = topDown ? y : height - 1 - y;
                int rowOffset = pixelOffset + (y * rowStride);
                for (int x = 0; x < width; x++) {
                    OfficeColor color = pixels[(sourceY * width) + x];
                    int offset = rowOffset + (x * 3);
                    bytes[offset] = color.B;
                    bytes[offset + 1] = color.G;
                    bytes[offset + 2] = color.R;
                }
            }

            return bytes;
        }

        private static byte[] CreateBmp32(int width, int height, IReadOnlyList<OfficeColor> pixels, bool topDown = false) {
            int rowStride = width * 4;
            int pixelOffset = 54;
            byte[] bytes = new byte[pixelOffset + (rowStride * height)];
            bytes[0] = (byte)'B';
            bytes[1] = (byte)'M';
            WriteInt32LittleEndian(bytes, 2, bytes.Length);
            WriteInt32LittleEndian(bytes, 10, pixelOffset);
            WriteInt32LittleEndian(bytes, 14, 40);
            WriteInt32LittleEndian(bytes, 18, width);
            WriteInt32LittleEndian(bytes, 22, topDown ? -height : height);
            WriteUInt16LittleEndian(bytes, 26, 1);
            WriteUInt16LittleEndian(bytes, 28, 32);

            for (int y = 0; y < height; y++) {
                int sourceY = topDown ? y : height - 1 - y;
                int rowOffset = pixelOffset + (y * rowStride);
                for (int x = 0; x < width; x++) {
                    OfficeColor color = pixels[(sourceY * width) + x];
                    int offset = rowOffset + (x * 4);
                    bytes[offset] = color.B;
                    bytes[offset + 1] = color.G;
                    bytes[offset + 2] = color.R;
                    bytes[offset + 3] = color.A;
                }
            }

            return bytes;
        }

        private static void WriteInt32LittleEndian(byte[] bytes, int offset, int value) {
            bytes[offset] = (byte)value;
            bytes[offset + 1] = (byte)(value >> 8);
            bytes[offset + 2] = (byte)(value >> 16);
            bytes[offset + 3] = (byte)(value >> 24);
        }

        private static void WriteUInt16LittleEndian(byte[] bytes, int offset, int value) {
            bytes[offset] = (byte)value;
            bytes[offset + 1] = (byte)(value >> 8);
        }

        private static int CountPixelsNear(OfficeRasterImage image, OfficeColor expected) =>
            Enumerable.Range(0, image.Width * image.Height)
                .Count(index => {
                    OfficeColor color = image.GetPixel(index % image.Width, index / image.Width);
                    return Math.Abs(color.R - expected.R) <= 8 &&
                        Math.Abs(color.G - expected.G) <= 8 &&
                        Math.Abs(color.B - expected.B) <= 8 &&
                        color.A > 0;
                });

        private static int CountPixelsNearAlpha(OfficeRasterImage image, OfficeColor expected, int tolerance, byte minimumAlpha, byte maximumAlpha) =>
            Enumerable.Range(0, image.Width * image.Height)
                .Count(index => {
                    OfficeColor color = image.GetPixel(index % image.Width, index / image.Width);
                    return Math.Abs(color.R - expected.R) <= tolerance &&
                        Math.Abs(color.G - expected.G) <= tolerance &&
                        Math.Abs(color.B - expected.B) <= tolerance &&
                        color.A >= minimumAlpha &&
                        color.A <= maximumAlpha;
                });

        private static double ExtractSvgRectHeight(string svg) {
            const string attribute = "height=\"";
            int start = svg.IndexOf(attribute, StringComparison.Ordinal);
            Assert.True(start >= 0, "Expected SVG output to contain a rectangle height.");
            start += attribute.Length;
            int end = svg.IndexOf('"', start);
            Assert.True(end > start, "Expected SVG rectangle height to be a valid number.");
            return double.Parse(svg.Substring(start, end - start), System.Globalization.CultureInfo.InvariantCulture);
        }

        private static int CountOccurrences(string value, string pattern) {
            int count = 0;
            int index = 0;
            while ((index = value.IndexOf(pattern, index, StringComparison.Ordinal)) >= 0) {
                count++;
                index += pattern.Length;
            }

            return count;
        }

        private static double ExtractFirstSvgFontSize(string svg) {
            const string attribute = "font-size=\"";
            int start = svg.IndexOf(attribute, StringComparison.Ordinal);
            Assert.True(start >= 0, "Expected SVG text output to include a font-size attribute.");
            start += attribute.Length;
            int end = svg.IndexOf('"', start);
            Assert.True(end > start, "Expected SVG text output to include a valid font-size value.");
            return double.Parse(svg.Substring(start, end - start), System.Globalization.CultureInfo.InvariantCulture);
        }

        private static void AssertColorNear(OfficeColor actual, OfficeColor expected) {
            Assert.True(
                Math.Abs(actual.R - expected.R) <= 8 &&
                Math.Abs(actual.G - expected.G) <= 8 &&
                Math.Abs(actual.B - expected.B) <= 8 &&
                actual.A > 0,
                $"Expected ARGB near {expected.A},{expected.R},{expected.G},{expected.B} but got {actual.A},{actual.R},{actual.G},{actual.B}.");
        }

        private static void AssertRasterImagesEqual(OfficeRasterImage expected, OfficeRasterImage actual) {
            Assert.Equal(expected.Width, actual.Width);
            Assert.Equal(expected.Height, actual.Height);
            for (int y = 0; y < expected.Height; y++) {
                for (int x = 0; x < expected.Width; x++) {
                    OfficeColor expectedPixel = expected.GetPixel(x, y);
                    OfficeColor actualPixel = actual.GetPixel(x, y);
                    Assert.True(
                        expectedPixel.Equals(actualPixel),
                        $"Pixel mismatch at {x},{y}. Expected {expectedPixel.A},{expectedPixel.R},{expectedPixel.G},{expectedPixel.B}; actual {actualPixel.A},{actualPixel.R},{actualPixel.G},{actualPixel.B}.");
                }
            }
        }

        private sealed class ExplosiveTuplePointList : IReadOnlyList<(double X, double Y)> {
            public int Count => 1_000_000;

            public (double X, double Y) this[int index] => throw new InvalidOperationException("Point materialization should have been skipped.");

            public IEnumerator<(double X, double Y)> GetEnumerator() => throw new InvalidOperationException("Point enumeration should have been skipped.");

            IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
        }

    }
}
