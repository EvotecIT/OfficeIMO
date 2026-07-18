using System;
using System.IO;
using System.IO.Compression;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Stencils;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioPngExport {
        [Fact]
        public void DocumentCanExportFirstPageToNativePng() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Diagram").Size(6, 4);
            VisioShape start = page.AddRectangle(1, 2, 1.5, 0.75, "Start");
            start.FillColor = OfficeColor.FromRgb(238, 247, 255);
            start.LineColor = OfficeColor.FromRgb(37, 99, 235);
            start.TextStyle = new VisioTextStyle {
                Size = 12,
                Bold = true,
                Color = OfficeColor.FromRgb(17, 24, 39)
            };

            VisioShape decision = page.AddDiamond(4, 2, 1.2, 1.2, "OK?");
            VisioConnector connector = page.AddConnector(start, decision, ConnectorKind.RightAngle, VisioSide.Right, VisioSide.Left);
            connector.EndArrow = EndArrow.Arrow;
            connector.Label = "yes";
            connector.LabelPlacement = VisioConnectorLabelPlacement.Along(0.5, 0, 0.2);

            byte[] png = document.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White
            });

            AssertPngHeader(png, 600, 400);
            using MemoryStream blankStream = new();
            VisioDocument blank = VisioDocument.Create(blankStream);
            blank.AddPage("Blank").Size(6, 4);
            byte[] blankPng = blank.ToPng(new VisioPngSaveOptions { PixelsPerInch = 100, BackgroundColor = OfficeColor.White });
            Assert.NotEqual(blankPng, png);
        }

        [Fact]
        public void PageCanSaveNativePngToFileAndStream() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Export").Size(2, 1);
            page.AddEllipse(1, 0.5, 1, 0.5, "Node");

            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".png");
            try {
                page.SaveAsPng(path, new VisioPngSaveOptions { PixelsPerInch = 96 });
                byte[] fileBytes = File.ReadAllBytes(path);
                AssertPngHeader(fileBytes, 192, 96);

                using MemoryStream stream = new();
                document.SaveAsPng(stream, new VisioPngSaveOptions { PixelsPerInch = 96 });
                AssertPngHeader(stream.ToArray(), 192, 96);
            } finally {
                if (File.Exists(path)) {
                    File.Delete(path);
                }
            }
        }

        [Fact]
        public void RetainedPngApiUsesTheSharedRasterSafetyPlanner() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Bounded").Size(100, 100);
            page.AddRectangle(50, 50, 80, 80, "Bounded");
            using var output = new MemoryStream();

            OfficeImageExportResult result = page.SaveAsPng(
                output,
                new VisioPngSaveOptions {
                    PixelsPerInch = 96D,
                    Supersampling = 1,
                    MaximumRasterPixels = 10_000L
                });

            Assert.True((long)result.Width * result.Height <= 10_000L);
            Assert.Contains(
                result.Diagnostics,
                diagnostic => diagnostic.Code ==
                              OfficeImageExportDiagnosticCodes.RasterScaleReduced);
            Assert.InRange(result.DpiX, 0.9D, 1.1D);
            Assert.Equal(result.Bytes, output.ToArray());
        }

        [Fact]
        public void PngRendererDrawsStyledShapeTextBackgrounds() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Text Background").Size(3, 2);
            VisioShape shape = page.AddRectangle(1.5, 1, 1.6, 0.7, "Escalation");
            shape.FillPattern = 0;
            shape.LinePattern = 0;
            shape.TextStyle = new VisioTextStyle {
                Size = 11,
                TextWidth = 1.2,
                TextHeight = 0.42,
                BackgroundColor = OfficeColor.FromRgb(255, 0, 0),
                BackgroundTransparency = 0,
                Color = OfficeColor.Black
            };

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            int redPixels = 0;
            for (int i = 0; i < image.Pixels.Length; i += 4) {
                if (image.Pixels[i] > 220 && image.Pixels[i + 1] < 80 && image.Pixels[i + 2] < 80 && image.Pixels[i + 3] > 200) {
                    redPixels++;
                }
            }

            Assert.True(redPixels > 100, "Expected visible red text background pixels in the native PNG preview.");
        }

        [Fact]
        public void PngRendererHonorsTextLocPinOffsetsForShapeTextBackgrounds() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Text LocPin Background").Size(3, 2);
            VisioShape shape = page.AddRectangle(1.5, 1, 2, 1, "Offset");
            shape.FillPattern = 0;
            shape.LinePattern = 0;
            shape.TextStyle = new VisioTextStyle {
                Size = 10,
                TextPinX = 0.2,
                TextPinY = 0.2,
                TextWidth = 0.8,
                TextHeight = 0.4,
                TextLocPinX = 0,
                TextLocPinY = 0,
                BackgroundColor = OfficeColor.FromRgb(255, 0, 0),
                BackgroundTransparency = 0,
                Color = OfficeColor.Transparent
            };

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            Assert.True(IsRedPixel(image, 110, 110), "Expected text background center to account for TxtLocPin offsets.");
            Assert.True(IsWhitePixel(image, 70, 130), "Expected the old TxtPin-only center to remain untouched.");
        }

        [Fact]
        public void PngRendererWrapsLongWordsInsideTextBounds() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Long Word Wrap").Size(3, 2);
            VisioShape shape = page.AddRectangle(1.5, 1, 1.2, 1.7, "MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM");
            shape.FillPattern = 0;
            shape.LinePattern = 0;
            shape.TextStyle = new VisioTextStyle {
                Size = 18,
                TextWidth = 0.32,
                TextHeight = 1.55,
                BackgroundColor = OfficeColor.FromRgb(255, 0, 0),
                BackgroundTransparency = 0,
                Color = OfficeColor.Transparent
            };

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            int minY = image.Height;
            int maxY = 0;
            for (int y = 0; y < image.Height; y++) {
                for (int x = 0; x < image.Width; x++) {
                    if (IsRedPixel(image, x, y)) {
                        minY = Math.Min(minY, y);
                        maxY = Math.Max(maxY, y);
                    }
                }
            }

            int redSpan = maxY - minY;
            Assert.True(redSpan > 40, "Expected a long unspaced word to wrap into multiple native PNG text-background lines instead of being shrunk to one line. Actual red background span: " + redSpan + ".");
        }

        [Fact]
        public void PngRendererBlendsStyledTextOpacity() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Text Opacity").Size(3, 2);
            VisioShape shape = page.AddRectangle(1.5, 1, 1.8, 0.8, "OfficeIMO");
            shape.FillPattern = 0;
            shape.LinePattern = 0;
            shape.TextStyle = new VisioTextStyle {
                Size = 24,
                TextWidth = 1.6,
                TextHeight = 0.55,
                Color = OfficeColor.FromRgba(220, 38, 38, 128)
            };

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            int translucentRedPixels = 0;
            for (int i = 0; i < image.Pixels.Length; i += 4) {
                if (image.Pixels[i] > 210 &&
                    image.Pixels[i + 1] > 95 && image.Pixels[i + 1] < 180 &&
                    image.Pixels[i + 2] > 95 && image.Pixels[i + 2] < 180 &&
                    image.Pixels[i + 3] > 200) {
                    translucentRedPixels++;
                }
            }

            Assert.True(translucentRedPixels > 60, "Expected semi-transparent text pixels to blend with the native PNG background.");
        }

        [Fact]
        public void PngRendererDrawsStyledTextUnderline() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Text Underline").Size(3, 2);
            VisioShape shape = page.AddRectangle(1.5, 1, 1.8, 0.8, "iiiiiiii");
            shape.FillPattern = 0;
            shape.LinePattern = 0;
            shape.TextStyle = new VisioTextStyle {
                Size = 24,
                TextWidth = 1.6,
                TextHeight = 0.55,
                Underline = true,
                Color = OfficeColor.FromRgb(22, 101, 52)
            };

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            int longestGreenRun = 0;
            for (int y = 105; y <= 125; y++) {
                int run = 0;
                for (int x = 80; x <= 220; x++) {
                    if (IsGreenPixel(image, x, y)) {
                        run++;
                        longestGreenRun = Math.Max(longestGreenRun, run);
                    } else {
                        run = 0;
                    }
                }
            }

            Assert.True(longestGreenRun > 35, "Expected a continuous styled underline in the native PNG render.");
        }

        [Fact]
        public void PngRendererDrawsStyledTextItalic() {
            RgbaPng upright = DecodeRgbaPng(RenderStyledItalicText(false));
            RgbaPng italic = DecodeRgbaPng(RenderStyledItalicText(true));

            Assert.Equal(upright.Width, italic.Width);
            Assert.Equal(upright.Height, italic.Height);

            int changedPixels = 0;
            for (int i = 0; i < upright.Pixels.Length; i += 4) {
                int delta = Math.Abs(upright.Pixels[i] - italic.Pixels[i]) +
                    Math.Abs(upright.Pixels[i + 1] - italic.Pixels[i + 1]) +
                    Math.Abs(upright.Pixels[i + 2] - italic.Pixels[i + 2]) +
                    Math.Abs(upright.Pixels[i + 3] - italic.Pixels[i + 3]);
                if (delta > 80) {
                    changedPixels++;
                }
            }

            Assert.True(changedPixels > 50, "Expected styled italic text to alter the native PNG text outline.");
        }

        [Fact]
        public void PngRendererRotatesStyledTextWithTextAngle() {
            RgbaPng upright = DecodeRgbaPng(RenderStyledRotatedText(0D));
            RgbaPng rotated = DecodeRgbaPng(RenderStyledRotatedText(Math.PI / 5D));

            Assert.Equal(upright.Width, rotated.Width);
            Assert.Equal(upright.Height, rotated.Height);

            int changedPixels = 0;
            for (int i = 0; i < upright.Pixels.Length; i += 4) {
                int delta = Math.Abs(upright.Pixels[i] - rotated.Pixels[i]) +
                    Math.Abs(upright.Pixels[i + 1] - rotated.Pixels[i + 1]) +
                    Math.Abs(upright.Pixels[i + 2] - rotated.Pixels[i + 2]) +
                    Math.Abs(upright.Pixels[i + 3] - rotated.Pixels[i + 3]);
                if (delta > 80) {
                    changedPixels++;
                }
            }

            Assert.True(changedPixels > 80, "Expected TextAngle to rotate the native PNG text outline.");
        }

        [Fact]
        public void PngRendererRotatesStyledTextBackgroundWithTextAngle() {
            RgbaPng upright = DecodeRgbaPng(RenderStyledRotatedTextBackground(0D));
            RgbaPng rotated = DecodeRgbaPng(RenderStyledRotatedTextBackground(Math.PI / 5D));

            Assert.Equal(upright.Width, rotated.Width);
            Assert.Equal(upright.Height, rotated.Height);

            int changedPixels = 0;
            for (int i = 0; i < upright.Pixels.Length; i += 4) {
                int delta = Math.Abs(upright.Pixels[i] - rotated.Pixels[i]) +
                    Math.Abs(upright.Pixels[i + 1] - rotated.Pixels[i + 1]) +
                    Math.Abs(upright.Pixels[i + 2] - rotated.Pixels[i + 2]) +
                    Math.Abs(upright.Pixels[i + 3] - rotated.Pixels[i + 3]);
                if (delta > 80) {
                    changedPixels++;
                }
            }

            Assert.True(changedPixels > 120, "Expected TextAngle to rotate the native PNG styled text background.");
        }

        [Fact]
        public void PngRendererPreservesDashedShapeOutlines() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Dashed Shape").Size(3, 2);
            VisioShape shape = page.AddRectangle(1.5, 1, 1, 1, string.Empty);
            shape.FillPattern = 0;
            shape.LineColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LinePattern = 2;
            shape.LineWeight = 0.04D;

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            bool sawDash = false;
            bool sawGap = false;
            for (int x = 110; x <= 190; x++) {
                sawDash |= IsRedPixel(image, x, 50);
                sawGap |= IsWhitePixel(image, x, 50);
            }

            Assert.True(sawDash, "Expected a painted dash on the top border.");
            Assert.True(sawGap, "Expected the top border dash gap to remain unpainted.");
        }

        [Fact]
        public void PngRendererSuppressesZeroWeightShapeOutlines() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Zero Line").Size(3, 2);
            VisioShape shape = page.AddRectangle(1.5, 1, 1, 1, string.Empty);
            shape.FillPattern = 0;
            shape.LineColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineWeight = 0D;

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            int redPixels = 0;
            for (int i = 0; i < image.Pixels.Length; i += 4) {
                if (image.Pixels[i] > 180 &&
                    image.Pixels[i + 1] < 90 &&
                    image.Pixels[i + 2] < 90 &&
                    image.Pixels[i + 3] > 200) {
                    redPixels++;
                }
            }

            Assert.Equal(0, redPixels);
        }

        [Fact]
        public void PngRendererDrawsSemanticDatabaseShapesAsCylinders() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Database Shape").Size(3, 2);
            VisioShape shape = page.AddRectangle(1.5, 1, 1, 1, string.Empty);
            shape.NameU = "Data";
            shape.FillColor = OfficeColor.FromRgb(37, 99, 235);
            shape.LineColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineWeight = 0.03D;
            shape.SetUserCell("OfficeIMO.StencilId", "architecture.database", "STR");

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            bool sawCapSeam = false;
            for (int y = 80; y <= 84; y++) {
                for (int x = 145; x <= 155; x++) {
                    sawCapSeam |= IsRedPixel(image, x, y);
                }
            }

            Assert.True(IsBluePixel(image, 150, 100), "Expected semantic database cylinder body fill in the native PNG render.");
            Assert.True(sawCapSeam, "Expected semantic database cylinder cap seam in the native PNG render.");
        }

        [Fact]
        public void PngRendererDrawsChevronShapesAsChevronPolygons() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Chevron Shape").Size(3, 2);
            VisioShape shape = page.AddRectangle(1.5, 1, 1, 1, string.Empty);
            shape.NameU = "Chevron";
            shape.FillColor = OfficeColor.FromRgb(37, 99, 235);
            shape.LinePattern = 0;

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            Assert.True(IsWhitePixel(image, 108, 100), "Expected the chevron notch to leave the former rectangle side untouched.");
            Assert.True(IsBluePixel(image, 150, 100), "Expected the chevron body fill in the native PNG render.");
            Assert.True(IsBluePixel(image, 190, 100), "Expected the chevron point fill in the native PNG render.");
        }

        [Fact]
        public void PngRendererDrawsFlowchartStartEndStencilsAsTerminatorCapsules() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Terminator Shape").Size(4, 2);
            VisioShape shape = page.AddStencilShape(VisioStencils.Flowchart, "flow.start-end", "start", 2, 1, 1.6, 0.8, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(37, 99, 235);
            shape.LinePattern = 0;

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            Assert.True(IsWhitePixel(image, 125, 65), "Expected the rounded terminator corner to leave the former rectangle corner untouched.");
            Assert.True(IsBluePixel(image, 160, 65), "Expected the terminator capsule top shoulder to render instead of collapsing to an ellipse.");
            Assert.True(IsBluePixel(image, 200, 100), "Expected the terminator capsule body fill in the native PNG render.");
        }

        [Fact]
        public void PngRendererDrawsDocumentStencilsAsWavyDocuments() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Document Shape").Size(3, 2);
            VisioShape shape = page.AddStencilShape(VisioStencils.CollaborationBusiness, "collab.document", "doc", 1.5, 1, 1.4, 1, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(37, 99, 235);
            shape.LinePattern = 0;

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            Assert.True(IsWhitePixel(image, 85, 145), "Expected the wavy document bottom to leave the former rectangle corner untouched.");
            Assert.True(IsBluePixel(image, 150, 100), "Expected the document body fill in the native PNG render.");
            Assert.True(HasBluePixelNear(image, 160, 128, radius: 2), "Expected the document wave crest to render in the native PNG render.");
        }

        [Fact]
        public void PngRendererDrawsDelayShapesAsDShapes() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Delay Shape").Size(3, 2);
            VisioShape shape = page.AddRectangle(1.5, 1, 1.4, 1, string.Empty);
            shape.NameU = "Delay";
            shape.FillColor = OfficeColor.FromRgb(37, 99, 235);
            shape.LinePattern = 0;

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            Assert.True(IsWhitePixel(image, 216, 55), "Expected the rounded delay corner to leave the former rectangle corner untouched.");
            Assert.True(IsBluePixel(image, 215, 100), "Expected the delay D-shape rounded side fill in the native PNG render.");
            Assert.True(IsBluePixel(image, 100, 100), "Expected the delay D-shape body fill in the native PNG render.");
        }

        [Fact]
        public void PngRendererDrawsManualInputShapesAsSlantedQuadrilaterals() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Manual Input Shape").Size(3, 2);
            VisioShape shape = page.AddRectangle(1.5, 1, 1, 1, string.Empty);
            shape.NameU = "Manual Input";
            shape.FillColor = OfficeColor.FromRgb(37, 99, 235);
            shape.LinePattern = 0;

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            Assert.True(IsWhitePixel(image, 195, 55), "Expected the slanted manual-input top to leave the former rectangle corner untouched.");
            Assert.True(IsBluePixel(image, 195, 80), "Expected the slanted manual-input top edge to include the upper-right shoulder.");
            Assert.True(IsBluePixel(image, 150, 100), "Expected the manual-input body fill in the native PNG render.");
        }

        [Fact]
        public void PngRendererRotatesEllipseShapesWithAngle() {
            RgbaPng upright = DecodeRgbaPng(RenderEllipseShape(0D));
            RgbaPng rotated = DecodeRgbaPng(RenderEllipseShape(Math.PI / 4D));

            Assert.Equal(upright.Width, rotated.Width);
            Assert.Equal(upright.Height, rotated.Height);

            int changedPixels = 0;
            for (int i = 0; i < upright.Pixels.Length; i += 4) {
                int delta = Math.Abs(upright.Pixels[i] - rotated.Pixels[i]) +
                    Math.Abs(upright.Pixels[i + 1] - rotated.Pixels[i + 1]) +
                    Math.Abs(upright.Pixels[i + 2] - rotated.Pixels[i + 2]) +
                    Math.Abs(upright.Pixels[i + 3] - rotated.Pixels[i + 3]);
                if (delta > 80) {
                    changedPixels++;
                }
            }

            Assert.True(changedPixels > 300, "Expected shape.Angle to rotate non-circular ellipse geometry in the native PNG render.");
        }

        [Fact]
        public void PngRendererPreservesVisioShapeAngleDirection() {
            RgbaPng image = DecodeRgbaPng(RenderEllipseShape(Math.PI / 4D));

            Assert.True(HasRedPixelNear(image, 105, 55, radius: 3), "Expected positive Visio shape.Angle to rotate the ellipse toward the upper-left.");
            Assert.True(HasRedPixelNear(image, 195, 145, radius: 3), "Expected positive Visio shape.Angle to rotate the ellipse toward the lower-right.");
            Assert.True(IsWhitePixel(image, 195, 55), "Expected the opposite upper-right diagonal to remain background.");
            Assert.True(IsWhitePixel(image, 105, 145), "Expected the opposite lower-left diagonal to remain background.");
        }

        [Fact]
        public void PngRendererCentersEllipseOnGeometryCenterWhenLocPinIsOffset() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Off Pin Ellipse").Size(3, 2);
            VisioShape shape = page.AddEllipse(1, 1, 1, 0.5, string.Empty);
            shape.LocPinX = 0;
            shape.LocPinY = 0;
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LinePattern = 0;

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            Assert.True(IsRedPixel(image, 150, 75), "Expected off-pin ellipse fill around the local geometry center.");
            Assert.True(IsWhitePixel(image, 100, 100), "Expected the shape pin to remain outside the off-pin ellipse geometry.");
        }

        [Fact]
        public void PngRendererPreservesRgbWhenDownsamplingTransparentEdges() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Transparent Edge").Size(2, 2);
            VisioShape shape = page.AddEllipse(1, 1, 1.1, 1.1, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LinePattern = 0;

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 80,
                BackgroundColor = OfficeColor.Transparent,
                Supersampling = 4
            });

            RgbaPng image = DecodeRgbaPng(png);
            int edgePixels = 0;
            for (int i = 0; i < image.Pixels.Length; i += 4) {
                byte alpha = image.Pixels[i + 3];
                if (alpha > 16 && alpha < 240) {
                    edgePixels++;
                    Assert.True(
                        image.Pixels[i] > 190 && image.Pixels[i + 1] < 80 && image.Pixels[i + 2] < 80,
                        "Expected transparent antialias edge pixels to retain the source red RGB, not average toward transparent black.");
                }
            }

            Assert.True(edgePixels > 12, "Expected transparent antialias edge pixels in the native PNG render.");
        }

        [Fact]
        public void PngRendererSuppressesArrowheadsWhenConnectorLineIsHidden() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Hidden Arrow").Size(4, 2);
            VisioShape source = page.AddRectangle(0.8, 1, 0.4, 0.4, string.Empty);
            source.FillPattern = 0;
            source.LinePattern = 0;
            VisioShape target = page.AddRectangle(3.2, 1, 0.4, 0.4, string.Empty);
            target.FillPattern = 0;
            target.LinePattern = 0;
            VisioConnector connector = page.AddConnector(source, target, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left);
            connector.BeginArrow = EndArrow.Arrow;
            connector.EndArrow = EndArrow.Arrow;
            connector.LineColor = OfficeColor.FromRgb(220, 38, 38);
            connector.LinePattern = 0;

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            int redPixels = 0;
            for (int i = 0; i < image.Pixels.Length; i += 4) {
                if (image.Pixels[i] > 180 &&
                    image.Pixels[i + 1] < 90 &&
                    image.Pixels[i + 2] < 90 &&
                    image.Pixels[i + 3] > 200) {
                    redPixels++;
                }
            }

            Assert.Equal(0, redPixels);
        }

        [Fact]
        public void PngRendererUsesFirstNonCollapsedSegmentForBeginArrowheads() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Collapsed Begin Arrow").Size(4, 2);
            VisioShape source = page.AddRectangle(1, 1, 1, 0.4, string.Empty);
            source.FillPattern = 0;
            source.LinePattern = 0;
            VisioShape target = page.AddRectangle(3, 1, 1, 0.4, string.Empty);
            target.FillPattern = 0;
            target.LinePattern = 0;
            VisioConnector connector = page.AddConnector(source, target, ConnectorKind.RightAngle, VisioSide.Right, VisioSide.Left);
            connector.BeginArrow = EndArrow.Arrow;
            connector.LineColor = OfficeColor.FromRgb(220, 38, 38);
            connector.LineWeight = 0.03;

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            int backwardArrowPixels = 0;
            for (int y = 88; y <= 112; y++) {
                for (int x = 125; x <= 145; x++) {
                    if (IsRedPixel(image, x, y)) {
                        backwardArrowPixels++;
                    }
                }
            }

            Assert.Equal(0, backwardArrowPixels);
        }

        [Fact]
        public void PngRendererCanUseConfiguredTrueTypeFontPath() {
            OfficeTrueTypeFont? font = OfficeTrueTypeFont.TryLoadDefault(out string? fontPath);
            if (font == null || string.IsNullOrWhiteSpace(fontPath)) {
                return;
            }

            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Configured Font").Size(3, 2);
            VisioShape shape = page.AddRectangle(1.5, 1, 2.2, 0.8, "OfficeIMO");
            shape.FillPattern = 0;
            shape.LinePattern = 0;
            shape.TextStyle = new VisioTextStyle {
                Size = 22,
                TextWidth = 2.0,
                TextHeight = 0.5,
                Color = OfficeColor.Black
            };

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                FontFilePath = fontPath,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            int darkPixels = 0;
            for (int i = 0; i < image.Pixels.Length; i += 4) {
                if (image.Pixels[i] < 80 && image.Pixels[i + 1] < 80 && image.Pixels[i + 2] < 80 && image.Pixels[i + 3] > 200) {
                    darkPixels++;
                }
            }

            Assert.True(darkPixels > 200, "Expected configured managed TrueType/OpenType font outlines in the native PNG render.");
        }

        [Fact]
        public void PngRendererNudgesConnectorLabelsAwayFromShapeCollisions() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Label Avoidance").Size(6, 3);
            VisioShape source = page.AddRectangle(1, 1.5, 1, 0.6, "Source");
            VisioShape target = page.AddRectangle(5, 1.5, 1, 0.6, "Target");
            VisioShape obstacle = page.AddRectangle(3, 1.5, 1.1, 0.7, string.Empty);
            obstacle.FillColor = OfficeColor.FromRgb(40, 96, 180);
            obstacle.LinePattern = 0;

            VisioConnector connector = page.AddConnector(source, target, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left);
            connector.LinePattern = 0;
            connector.Label = "handoff";
            connector.PlaceLabel(0.5, width: 1.2, height: 0.3);
            connector.TextStyle = new VisioTextStyle {
                BackgroundColor = OfficeColor.FromRgb(255, 0, 0),
                BackgroundTransparency = 0,
                Color = OfficeColor.Black
            };

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            int center = (((image.Height - 150) * image.Width) + 300) * 4;
            Assert.True(image.Pixels[center] < 100 && image.Pixels[center + 1] < 130 && image.Pixels[center + 2] > 150,
                "Expected the obstacle center to remain visible instead of being covered by the connector label background.");
        }

        [Fact]
        public void PngRendererNudgesConnectorLabelsAwayFromEndpointShapeCollisions() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Endpoint Label Avoidance").Size(6, 3);
            VisioShape source = page.AddRectangle(1, 1.5, 1, 0.6, string.Empty);
            source.FillColor = OfficeColor.FromRgb(40, 96, 180);
            source.LinePattern = 0;
            VisioShape target = page.AddRectangle(5, 1.5, 1, 0.6, string.Empty);
            target.LinePattern = 0;

            VisioConnector connector = page.AddConnector(source, target, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left);
            connector.LinePattern = 0;
            connector.Label = "endpoint collision";
            connector.PlaceLabel(0, width: 1.2, height: 0.3);
            connector.TextStyle = new VisioTextStyle {
                BackgroundColor = OfficeColor.FromRgb(255, 0, 0),
                BackgroundTransparency = 0,
                Color = OfficeColor.Black
            };

            VisioRenderConnectorLabelPlacement placement = VisioRenderLabelLayout.Create(page).Resolve(
                connector,
                new[] { (1.5D, 1.5D), (4.5D, 1.5D) });
            Assert.True(placement.Adjusted, "Expected endpoint collision avoidance to adjust the connector label placement.");
            Assert.True(placement.X > 2D, "Expected endpoint collision avoidance to move the label away from the source shape.");

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            AssertPngHeader(png, 600, 300);
        }

        [Fact]
        public void PngRendererHonorsConnectorLabelTextLocPin() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Connector Label LocPin").Size(4, 3);
            VisioShape source = page.AddRectangle(1, 1.5, 0.4, 0.4, string.Empty);
            VisioShape target = page.AddRectangle(3, 1.5, 0.4, 0.4, string.Empty);
            VisioConnector connector = page.AddConnector(source, target, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left);
            connector.LinePattern = 0;
            connector.Label = "locpin";
            connector.PlaceLabelAt(2, 1.5, width: 0.9, height: 0.4);
            connector.TextStyle = new VisioTextStyle {
                TextWidth = 0.9,
                TextHeight = 0.4,
                TextLocPinX = 0,
                TextLocPinY = 0,
                BackgroundColor = OfficeColor.FromRgb(255, 0, 0),
                BackgroundTransparency = 0,
                Color = OfficeColor.Transparent
            };

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            Assert.True(IsWhitePixel(image, 200, 150), "Expected the text pin itself to remain outside the rendered label background when TxtLocPin is bottom-left.");
            Assert.True(IsRedPixel(image, 245, 130), "Expected the connector label background to render around the TxtLocPin-adjusted text box center.");
        }

        [Fact]
        public void PngRendererNudgesConnectorLabelsAwayFromOtherConnectorLines() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Connector Label Crossing Avoidance").Size(6, 3);
            VisioShape source = page.AddRectangle(1, 1.5, 1, 0.4, string.Empty);
            VisioShape target = page.AddRectangle(5, 1.5, 1, 0.4, string.Empty);
            VisioConnector labeled = page.AddConnector(source, target, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left);
            labeled.Label = "handoff";
            labeled.PlaceLabel(0.5, width: 1.2, height: 0.3);
            labeled.TextStyle = new VisioTextStyle {
                BackgroundColor = OfficeColor.FromRgb(255, 0, 0),
                BackgroundTransparency = 0,
                Color = OfficeColor.Black
            };

            VisioShape top = page.AddRectangle(3, 2.7, 0.5, 0.4, string.Empty);
            VisioShape bottom = page.AddRectangle(3, 0.3, 0.5, 0.4, string.Empty);
            page.AddConnector(top, bottom, ConnectorKind.Straight, VisioSide.Bottom, VisioSide.Top);

            VisioRenderConnectorLabelPlacement placement = VisioRenderLabelLayout.Create(page).Resolve(
                labeled,
                new[] { (1.5D, 1.5D), (4.5D, 1.5D) });
            Assert.True(placement.Adjusted, "Expected connector-line collision avoidance to adjust the connector label placement.");
            Assert.True(Math.Abs(placement.X - 3D) > 0.6D, "Expected connector-line collision avoidance to move the label away from the crossing connector.");

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            AssertPngHeader(png, 600, 300);
        }

        [Fact]
        public void PngRendererKeepsDenseConnectorLabelsSeparated() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Dense Label Clearance").Size(6, 3);
            VisioShape lowerSource = page.AddRectangle(1, 1.5, 0.6, 0.25, string.Empty);
            VisioShape lowerTarget = page.AddRectangle(5, 1.5, 0.6, 0.25, string.Empty);
            VisioConnector lower = page.AddConnector(lowerSource, lowerTarget, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left);
            lower.Label = "phase one";
            lower.PlaceLabel(0.5, width: 1.2, height: 0.3);

            VisioShape upperSource = page.AddRectangle(1, 1.82, 0.6, 0.25, string.Empty);
            VisioShape upperTarget = page.AddRectangle(5, 1.82, 0.6, 0.25, string.Empty);
            VisioConnector upper = page.AddConnector(upperSource, upperTarget, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left);
            upper.Label = "phase two";
            upper.PlaceLabel(0.5, width: 1.2, height: 0.3);

            VisioRenderLabelLayout layout = VisioRenderLabelLayout.Create(page);
            VisioRenderConnectorLabelPlacement lowerPlacement = layout.Resolve(
                lower,
                new[] { (1.3D, 1.5D), (4.7D, 1.5D) });
            VisioRenderConnectorLabelPlacement upperPlacement = layout.Resolve(
                upper,
                new[] { (1.3D, 1.82D), (4.7D, 1.82D) });

            Assert.False(lowerPlacement.Adjusted);
            Assert.True(upperPlacement.Adjusted, "Expected dense label clearance to move the second label away from the first label.");
            Assert.True(Math.Abs(upperPlacement.Y - lowerPlacement.Y) > 0.32D, "Expected dense labels to keep a readable vertical gap.");

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            AssertPngHeader(png, 600, 300);
        }

        [Fact]
        public void PngRendererProjectsBuiltInStencilMetadataAsVectorArtwork() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Stencil Artwork").Size(3, 2);
            VisioShape shape = page.AddStencilShape(VisioStencils.SecurityIdentity, "sec.firewall", "firewall", 1.5, 1, 1.2, 0.8, string.Empty);
            shape.FillPattern = 0;
            shape.LinePattern = 0;

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            int nonWhitePixels = 0;
            for (int i = 0; i < image.Pixels.Length; i += 4) {
                if (image.Pixels[i] < 245 || image.Pixels[i + 1] < 245 || image.Pixels[i + 2] < 245) {
                    nonWhitePixels++;
                }
            }

            Assert.True(nonWhitePixels > 50, "Expected native PNG stencil artwork pixels without relying on shape fill, stroke, or text.");
        }

        [Fact]
        public void PngRendererRotatesStencilMetadataArtworkWithShape() {
            RgbaPng upright = DecodeRgbaPng(RenderStencilArtwork(0D));
            RgbaPng rotated = DecodeRgbaPng(RenderStencilArtwork(Math.PI / 4D));

            Assert.Equal(upright.Width, rotated.Width);
            Assert.Equal(upright.Height, rotated.Height);

            int changedPixels = 0;
            for (int i = 0; i < upright.Pixels.Length; i += 4) {
                int delta = Math.Abs(upright.Pixels[i] - rotated.Pixels[i]) +
                    Math.Abs(upright.Pixels[i + 1] - rotated.Pixels[i + 1]) +
                    Math.Abs(upright.Pixels[i + 2] - rotated.Pixels[i + 2]) +
                    Math.Abs(upright.Pixels[i + 3] - rotated.Pixels[i + 3]);
                if (delta > 80) {
                    changedPixels++;
                }
            }

            Assert.True(changedPixels > 60, "Expected rotated stencil pictograms to alter the native PNG render.");
        }

        [Fact]
        public void PngRendererRotatesCurvedStencilMetadataArtworkWithShape() {
            RgbaPng upright = DecodeRgbaPng(RenderStencilArtwork(0D, "database"));
            RgbaPng rotated = DecodeRgbaPng(RenderStencilArtwork(Math.PI / 2D, "database"));

            Assert.Equal(upright.Width, rotated.Width);
            Assert.Equal(upright.Height, rotated.Height);

            int changedPixels = 0;
            for (int i = 0; i < upright.Pixels.Length; i += 4) {
                int delta = Math.Abs(upright.Pixels[i] - rotated.Pixels[i]) +
                    Math.Abs(upright.Pixels[i + 1] - rotated.Pixels[i + 1]) +
                    Math.Abs(upright.Pixels[i + 2] - rotated.Pixels[i + 2]) +
                    Math.Abs(upright.Pixels[i + 3] - rotated.Pixels[i + 3]);
                if (delta > 80) {
                    changedPixels++;
                }
            }

            Assert.True(changedPixels > 100, "Expected curved stencil pictograms to rotate in the native PNG render.");
        }

        [Fact]
        public void PngRendererDoesNotProjectSequenceFragmentRegionsAsCloudArtwork() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Sequence Fragment").Size(5, 3);
            VisioShape fragment = page.AddRectangle(2.5, 1.5, 4, 2, string.Empty);
            fragment.FillPattern = 0;
            fragment.LinePattern = 0;
            fragment.SetUserCell("OfficeIMO.Kind", "SequenceFragment", "STR");
            fragment.SetUserCell("OfficeIMO.StencilId", "seq.fragment", "STR");
            fragment.SetUserCell("OfficeIMO.StencilName", "Combined Fragment", "STR");
            fragment.SetUserCell("OfficeIMO.StencilAliases", "alt;combined-fragment;critical;fragment;loop;opt;region", "STR");
            fragment.SetUserCell("OfficeIMO.StencilTags", "Rectangle;seq;Sequence Diagram", "STR");

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 48,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            int nonWhitePixels = 0;
            for (int i = 0; i < image.Pixels.Length; i += 4) {
                if (image.Pixels[i] < 245 || image.Pixels[i + 1] < 245 || image.Pixels[i + 2] < 245) {
                    nonWhitePixels++;
                }
            }

            Assert.Equal(0, nonWhitePixels);
        }

        [Fact]
        public void PngRendererProjectsPackageBackedPngPreviewArtwork() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Package Preview").Size(3, 2);
            AddPackagePreviewShape(page, TrueColorBluePng);

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            int bluePixels = 0;
            for (int i = 0; i < image.Pixels.Length; i += 4) {
                if (image.Pixels[i] < 60 && image.Pixels[i + 1] < 150 && image.Pixels[i + 2] > 180 && image.Pixels[i + 3] > 200) {
                    bluePixels++;
                }
            }

            Assert.True(bluePixels > 100, "Expected embedded package preview PNG pixels in the native PNG render.");
        }

        [Fact]
        public void PngRendererProjectsPackageBackedBmpPreviewArtwork() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Package BMP Preview").Size(3, 2);
            AddPackagePreviewShape(page, TrueColorBlueBmp, "image/bmp", ".bmp", "../media/image1.bmp");

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            int bluePixels = 0;
            for (int i = 0; i < image.Pixels.Length; i += 4) {
                if (image.Pixels[i] < 60 && image.Pixels[i + 1] < 150 && image.Pixels[i + 2] > 180 && image.Pixels[i + 3] > 200) {
                    bluePixels++;
                }
            }

            Assert.True(bluePixels > 100, "Expected embedded package preview BMP pixels in the native PNG render.");
        }

        [Fact]
        public void PngRendererProjectsPackageBackedSvgPreviewArtwork() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Package SVG Preview").Size(3, 2);
            string svg = "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 20 20\"><path fill=\"#0070c0\" fill-rule=\"evenodd\" d=\"M0 0 H20 V20 H0 Z M7 7 H13 V13 H7 Z\"/></svg>";
            AddPackagePreviewShape(page, Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(svg)), "image/svg+xml", ".svg", "../media/image1.svg");

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            int bluePixels = 0;
            for (int i = 0; i < image.Pixels.Length; i += 4) {
                if (image.Pixels[i] < 60 && image.Pixels[i + 1] < 150 && image.Pixels[i + 2] > 180 && image.Pixels[i + 3] > 200) {
                    bluePixels++;
                }
            }

            Assert.True(bluePixels > 100, "Expected embedded package preview SVG pixels in the native PNG render.");
            Assert.True(IsWhitePixel(image, 150, 100), "Expected compound SVG preview paths to preserve even-odd holes.");
        }

        [Fact]
        public void PngRendererProjectsPackageBackedSvgCssStyledPreviewArtwork() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Package SVG CSS Preview").Size(3, 2);
            string svg = "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 20 20\"><style>path{fill:#dc2626}#badge{fill:#0070c0}.cutout,path.cutout{fill-rule:evenodd}</style><path id=\"badge\" class=\"cutout\" d=\"M0 0 H20 V20 H0 Z M7 7 H13 V13 H7 Z\"/></svg>";
            AddPackagePreviewShape(page, Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(svg)), "image/svg+xml", ".svg", "../media/css.svg");

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            int bluePixels = 0;
            for (int i = 0; i < image.Pixels.Length; i += 4) {
                if (image.Pixels[i] < 60 && image.Pixels[i + 1] < 150 && image.Pixels[i + 2] > 180 && image.Pixels[i + 3] > 200) {
                    bluePixels++;
                }
            }

            Assert.True(bluePixels > 100, "Expected CSS id selectors in embedded SVG previews to override element selectors through the native PNG path.");
            Assert.True(IsWhitePixel(image, 150, 100), "Expected CSS class and element.class fill-rule declarations to preserve even-odd holes.");
        }

        [Fact]
        public void PngRendererProjectsPackageBackedSvgCurrentColorPreviewArtwork() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Package SVG CurrentColor Preview").Size(3, 2);
            string svg = "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 20 20\"><style>.accent{color:#0070c0}.mark{fill:currentColor}.stop{stop-color:currentColor}</style><defs><linearGradient id=\"fade\" x1=\"0%\" y1=\"0%\" x2=\"100%\" y2=\"0%\"><stop class=\"stop\" offset=\"0%\"/><stop offset=\"100%\" stop-color=\"#dc2626\"/></linearGradient></defs><g class=\"accent\"><rect class=\"mark\" x=\"0\" y=\"0\" width=\"10\" height=\"20\"/><rect x=\"10\" y=\"0\" width=\"10\" height=\"20\" fill=\"url(#fade)\"/></g></svg>";
            AddPackagePreviewShape(page, Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(svg)), "image/svg+xml", ".svg", "../media/current-color.svg");

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            int bluePixels = 0;
            int redPixels = 0;
            for (int i = 0; i < image.Pixels.Length; i += 4) {
                if (image.Pixels[i] < 80 && image.Pixels[i + 1] < 150 && image.Pixels[i + 2] > 160 && image.Pixels[i + 3] > 200) {
                    bluePixels++;
                }

                if (image.Pixels[i] > 170 && image.Pixels[i + 1] < 90 && image.Pixels[i + 2] < 90 && image.Pixels[i + 3] > 200) {
                    redPixels++;
                }
            }

            Assert.True(bluePixels > 100, $"Expected SVG currentColor fill and gradient stops in package preview output, but found {bluePixels} blue pixels.");
            Assert.True(redPixels > 50, $"Expected SVG currentColor gradient to keep its literal red stop in package preview output, but found {redPixels} red pixels.");
            Assert.True(IsBluePixel(image, 130, 100), "Expected currentColor fill to render on the left side of the package preview.");
        }

        [Fact]
        public void PngRendererProjectsPackageBackedSvgTextPreviewArtwork() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Package SVG Text Preview").Size(3, 2);
            string svg = "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 20 20\"><style>.label{fill:#0070c0;font-size:9px;font-weight:700;text-anchor:middle}</style><text class=\"label\" x=\"10\" y=\"13\"><tspan>IMO</tspan></text></svg>";
            AddPackagePreviewShape(page, Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(svg)), "image/svg+xml", ".svg", "../media/text-label.svg");

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            int bluePixels = 0;
            for (int i = 0; i < image.Pixels.Length; i += 4) {
                if (image.Pixels[i] < 90 && image.Pixels[i + 1] < 170 && image.Pixels[i + 2] > 140 && image.Pixels[i + 3] > 200) {
                    bluePixels++;
                }
            }

            Assert.True(bluePixels > 40, $"Expected SVG text labels in package preview output, but found {bluePixels} blue text pixels.");
            Assert.False(IsWhitePixel(image, 150, 100), "Expected centered SVG text label pixels in the package preview.");
        }

        [Fact]
        public void PngRendererPreservesPackageBackedSvgTextXmlSpace() {
            static RgbaPng RenderSvgText(string svg) {
                using MemoryStream packageStream = new();
                VisioDocument document = VisioDocument.Create(packageStream);
                VisioPage page = document.AddPage("Package SVG Text Space Preview").Size(3, 2);
                AddPackagePreviewShape(page, Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(svg)), "image/svg+xml", ".svg", "../media/text-space-label.svg");

                byte[] png = page.ToPng(new VisioPngSaveOptions {
                    PixelsPerInch = 100,
                    BackgroundColor = OfficeColor.White,
                    Supersampling = 1
                });

                return DecodeRgbaPng(png);
            }

            static (int Span, int Count) MeasureBlueSpan(RgbaPng image) {
                int minX = image.Width;
                int maxX = -1;
                int count = 0;
                for (int y = 0; y < image.Height; y++) {
                    for (int x = 0; x < image.Width; x++) {
                        if (!IsBluePixel(image, x, y)) {
                            continue;
                        }

                        minX = Math.Min(minX, x);
                        maxX = Math.Max(maxX, x);
                        count++;
                    }
                }

                return (maxX >= minX ? maxX - minX : 0, count);
            }

            const string preservedSvg = "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 60 20\" xml:space=\"preserve\"><style>.label{fill:#0070c0;font-size:12px;font-weight:700}</style><text class=\"label\" x=\"2\" y=\"14\"><tspan>A     B</tspan></text></svg>";
            const string collapsedSvg = "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 60 20\"><style>.label{fill:#0070c0;font-size:12px;font-weight:700}</style><text class=\"label\" x=\"2\" y=\"14\"><tspan>A     B</tspan></text></svg>";

            (int preservedSpan, int preservedCount) = MeasureBlueSpan(RenderSvgText(preservedSvg));
            (int collapsedSpan, int collapsedCount) = MeasureBlueSpan(RenderSvgText(collapsedSvg));

            Assert.True(preservedCount > 20, $"Expected preserved-space SVG text to render visible blue text pixels, but found {preservedCount}.");
            Assert.True(collapsedCount > 20, $"Expected collapsed-space SVG text to render visible blue text pixels, but found {collapsedCount}.");
            Assert.True(preservedSpan > collapsedSpan + 8, $"Expected xml:space='preserve' to keep a wider text span, but preserved={preservedSpan}px and collapsed={collapsedSpan}px.");
        }

        [Fact]
        public void PngRendererSkipsPackageBackedSvgHiddenPreviewArtwork() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Package SVG Hidden Preview").Size(3, 2);
            string svg = "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 20 20\"><style>.hidden{display:none}.ghost{visibility:hidden}</style><rect class=\"hidden\" x=\"0\" y=\"0\" width=\"20\" height=\"20\" fill=\"#dc2626\"/><g class=\"ghost\"><rect x=\"0\" y=\"0\" width=\"20\" height=\"20\" fill=\"#dc2626\"/></g><rect x=\"2\" y=\"2\" width=\"16\" height=\"16\" fill=\"#0070c0\"/></svg>";
            AddPackagePreviewShape(page, Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(svg)), "image/svg+xml", ".svg", "../media/hidden.svg");

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            int redPixels = 0;
            int bluePixels = 0;
            for (int i = 0; i < image.Pixels.Length; i += 4) {
                if (image.Pixels[i] > 170 && image.Pixels[i + 1] < 90 && image.Pixels[i + 2] < 90 && image.Pixels[i + 3] > 200) {
                    redPixels++;
                }

                if (image.Pixels[i] < 80 && image.Pixels[i + 1] < 150 && image.Pixels[i + 2] > 160 && image.Pixels[i + 3] > 200) {
                    bluePixels++;
                }
            }

            Assert.Equal(0, redPixels);
            Assert.True(bluePixels > 100, $"Expected visible SVG preview artwork to render while hidden layers are skipped, but found {bluePixels} blue pixels.");
            Assert.True(IsBluePixel(image, 150, 100), "Expected visible SVG preview artwork in the package preview center.");
        }

        [Fact]
        public void PngRendererProjectsPackageBackedSvgUseSymbolPreviewArtwork() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Package SVG Use Preview").Size(3, 2);
            string svg = "<svg xmlns=\"http://www.w3.org/2000/svg\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" viewBox=\"0 0 20 20\"><defs><symbol id=\"badge\" viewBox=\"0 0 10 10\"><path fill=\"#0070c0\" fill-rule=\"evenodd\" d=\"M0 0 H10 V10 H0 Z M4 4 H6 V6 H4 Z\"/></symbol></defs><use xlink:href=\"#badge\" width=\"20\" height=\"20\"/></svg>";
            AddPackagePreviewShape(page, Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(svg)), "image/svg+xml", ".svg", "../media/use-symbol.svg");

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            int bluePixels = 0;
            for (int i = 0; i < image.Pixels.Length; i += 4) {
                if (image.Pixels[i] < 60 && image.Pixels[i + 1] < 150 && image.Pixels[i + 2] > 180 && image.Pixels[i + 3] > 200) {
                    bluePixels++;
                }
            }

            Assert.True(bluePixels > 100, $"Expected package-backed SVG use/symbol artwork to render through the native PNG path, but found {bluePixels} blue pixels.");
            Assert.True(IsWhitePixel(image, 150, 100), "Expected use/symbol viewBox scaling to preserve the referenced even-odd center hole.");
        }

        [Fact]
        public void PngRendererProjectsPackageBackedSvgEmbeddedRasterPreviewArtwork() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Package SVG Embedded Raster Preview").Size(3, 2);
            string svg = $"<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 20 20\"><image href=\"data:image/png;base64,{TrueColorBluePng}\" x=\"4\" y=\"4\" width=\"12\" height=\"12\"/></svg>";
            AddPackagePreviewShape(page, Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(svg)), "image/svg+xml", ".svg", "../media/embedded-raster.svg");

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            int bluePixels = 0;
            for (int i = 0; i < image.Pixels.Length; i += 4) {
                if (image.Pixels[i] < 60 && image.Pixels[i + 1] < 150 && image.Pixels[i + 2] > 180 && image.Pixels[i + 3] > 200) {
                    bluePixels++;
                }
            }

            Assert.True(bluePixels > 100, $"Expected embedded raster image pixels inside SVG package previews to render through the native PNG path, but found {bluePixels} blue pixels.");
            Assert.True(IsBluePixel(image, 150, 100), "Expected embedded SVG raster artwork to land in the package preview center.");
        }

        [Fact]
        public void PngRendererProjectsPackageBackedSvgRelatedRasterPreviewArtwork() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Package SVG Related Raster Preview").Size(3, 2);
            string svg = "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 20 20\"><image href=\"related.png\" x=\"4\" y=\"4\" width=\"12\" height=\"12\"/></svg>";
            VisioShape shape = AddPackagePreviewShape(page, Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(svg)), "image/svg+xml", ".svg", "../media/preview.svg");
            shape.Master!.RawMasterRelationships.Add(new VisioAssets.MasterRelationshipContent {
                Id = "rIdRelatedImage",
                Type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                Target = "../media/related.png",
                ContentType = "image/png",
                Extension = ".png",
                Data = Convert.FromBase64String(TrueColorBluePng)
            });

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            int bluePixels = 0;
            for (int i = 0; i < image.Pixels.Length; i += 4) {
                if (image.Pixels[i] < 60 && image.Pixels[i + 1] < 150 && image.Pixels[i + 2] > 180 && image.Pixels[i + 3] > 200) {
                    bluePixels++;
                }
            }

            Assert.True(bluePixels > 100, $"Expected related package image pixels inside SVG previews to render through the native PNG path, but found {bluePixels} blue pixels.");
            Assert.True(IsBluePixel(image, 150, 100), "Expected related SVG raster artwork to land in the package preview center.");
        }

        [Fact]
        public void PngRendererProjectsPackageBackedSvgGradientPreviewArtwork() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Package SVG Gradient Preview").Size(3, 2);
            string svg = "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 20 20\"><defs><linearGradient id=\"accent\" x1=\"0%\" y1=\"0%\" x2=\"100%\" y2=\"0%\"><stop offset=\"0%\" stop-color=\"#0070c0\"/><stop offset=\"100%\" stop-color=\"#dc2626\"/></linearGradient></defs><g fill=\"url(#accent)\"><rect x=\"0\" y=\"0\" width=\"20\" height=\"20\"/></g></svg>";
            AddPackagePreviewShape(page, Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(svg)), "image/svg+xml", ".svg", "../media/gradient.svg");

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            int bluePixels = 0;
            int redPixels = 0;
            for (int i = 0; i < image.Pixels.Length; i += 4) {
                if (image.Pixels[i] < 80 && image.Pixels[i + 1] < 150 && image.Pixels[i + 2] > 160 && image.Pixels[i + 3] > 200) {
                    bluePixels++;
                }

                if (image.Pixels[i] > 170 && image.Pixels[i + 1] < 90 && image.Pixels[i + 2] < 90 && image.Pixels[i + 3] > 200) {
                    redPixels++;
                }
            }

            Assert.True(bluePixels > 50, $"Expected SVG linear gradient blue stop pixels in package preview output, but found {bluePixels}.");
            Assert.True(redPixels > 50, $"Expected SVG linear gradient red stop pixels in package preview output, but found {redPixels}.");
        }

        [Fact]
        public void PngRendererProjectsPackageBackedSvgClippedPreviewArtwork() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Package SVG Clipped Preview").Size(3, 2);
            string svg = "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 20 20\"><defs><clipPath id=\"rightHalf\"><rect x=\"10\" y=\"0\" width=\"10\" height=\"20\"/></clipPath></defs><rect x=\"0\" y=\"0\" width=\"20\" height=\"20\" fill=\"#0070c0\" clip-path=\"url(#rightHalf)\"/></svg>";
            AddPackagePreviewShape(page, Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(svg)), "image/svg+xml", ".svg", "../media/clip.svg");

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            Assert.True(IsWhitePixel(image, 135, 100), "Expected SVG clipPath to leave the left side of package preview artwork unpainted.");
            Assert.True(IsBluePixel(image, 165, 100), "Expected SVG clipPath to keep the right side of package preview artwork painted.");
        }

        [Fact]
        public void PngRendererProjectsPackageBackedSvgDashedStrokePreviewArtwork() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Package SVG Dashed Stroke Preview").Size(3, 2);
            string svg = "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 20 20\"><g stroke=\"#0070c0\" stroke-width=\"2\" stroke-dasharray=\"2 2\" fill=\"none\"><path d=\"M2 10 H18\"/></g></svg>";
            AddPackagePreviewShape(page, Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(svg)), "image/svg+xml", ".svg", "../media/dashed-stroke.svg");

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            bool sawDash = false;
            bool sawGap = false;
            for (int x = 130; x <= 170; x++) {
                sawDash |= IsBluePixel(image, x, 100);
                sawGap |= IsWhitePixel(image, x, 100);
            }

            Assert.True(sawDash, "Expected SVG stroke-dasharray to paint dash segments in package preview output.");
            Assert.True(sawGap, "Expected SVG stroke-dasharray to leave visible gaps in package preview output.");
        }

        [Fact]
        public void PngRendererProjectsPackageBackedSvgDashedGradientStrokePreviewArtwork() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Package SVG Dashed Gradient Stroke Preview").Size(3, 2);
            string svg = "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 20 20\"><defs><linearGradient id=\"accent\" x1=\"0%\" y1=\"0%\" x2=\"100%\" y2=\"0%\"><stop offset=\"0%\" stop-color=\"#0070c0\"/><stop offset=\"100%\" stop-color=\"#dc2626\"/></linearGradient></defs><g stroke=\"url(#accent)\" stroke-width=\"2\" stroke-dasharray=\"2 2\" fill=\"none\"><path d=\"M2 10 H18\"/></g></svg>";
            AddPackagePreviewShape(page, Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(svg)), "image/svg+xml", ".svg", "../media/dashed-gradient-stroke.svg");

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            bool sawDash = false;
            bool sawGap = false;
            for (int x = 130; x <= 170; x++) {
                sawDash |= IsRedPixel(image, x, 100);
                sawGap |= IsWhitePixel(image, x, 100);
            }

            Assert.True(sawDash, "Expected SVG dashed gradient stroke segments to paint in package preview output.");
            Assert.True(sawGap, "Expected SVG stroke-dasharray to leave visible gaps in package preview output.");
        }

        [Fact]
        public void PngRendererProjectsPackageBackedSvgRoundStrokeCapsPreviewArtwork() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Package SVG Round Stroke Cap Preview").Size(3, 2);
            string svg = "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 20 20\"><path d=\"M5 6 H15\" fill=\"none\" stroke=\"#dc2626\" stroke-width=\"4\" stroke-linecap=\"butt\"/><path d=\"M5 14 H15\" fill=\"none\" stroke=\"#0070c0\" stroke-width=\"4\" stroke-linecap=\"round\"/></svg>";
            AddPackagePreviewShape(page, Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(svg)), "image/svg+xml", ".svg", "../media/round-stroke-cap.svg");

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            (int RedMin, int RedMax, int RedCount) = FindHorizontalColorSpanNear(image, 84, radius: 2, IsRedPixel);
            (int BlueMin, int BlueMax, int BlueCount) = FindHorizontalColorSpanNear(image, 116, radius: 2, IsBluePixel);
            Assert.True(RedCount > 0, "Expected SVG stroke-linecap=\"butt\" comparison stroke to render in package preview output.");
            Assert.True(BlueCount > 0, "Expected SVG stroke-linecap=\"round\" comparison stroke to render in package preview output.");
            Assert.True(BlueMin < RedMin, $"Expected SVG stroke-linecap=\"round\" to extend before the butt cap start. Red min {RedMin}, blue min {BlueMin}.");
            Assert.True(BlueMax > RedMax, $"Expected SVG stroke-linecap=\"round\" to extend after the butt cap end. Red max {RedMax}, blue max {BlueMax}.");
        }

        [Fact]
        public void PngRendererProjectsPackageBackedSvgRoundStrokeJoinsPreviewArtwork() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Package SVG Round Stroke Join Preview").Size(3, 2);
            string svg = "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"256\" height=\"256\" viewBox=\"0 0 20 28\"><path d=\"M3 11 L10 3 L17 11\" fill=\"none\" stroke=\"#dc2626\" stroke-width=\"5\" stroke-linecap=\"butt\" stroke-linejoin=\"bevel\"/><path d=\"M3 25 L10 17 L17 25\" fill=\"none\" stroke=\"#0070c0\" stroke-width=\"5\" stroke-linecap=\"butt\" stroke-linejoin=\"round\"/></svg>";
            AddPackagePreviewShape(page, Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(svg)), "image/svg+xml", ".svg", "../media/round-stroke-join.svg");

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            int redPixels = 0;
            int bluePixels = 0;
            for (int i = 0; i < image.Pixels.Length; i += 4) {
                if (image.Pixels[i] > 180 && image.Pixels[i + 1] < 80 && image.Pixels[i + 2] < 80 && image.Pixels[i + 3] > 200) {
                    redPixels++;
                }

                if (image.Pixels[i] < 60 && image.Pixels[i + 1] < 150 && image.Pixels[i + 2] > 180 && image.Pixels[i + 3] > 200) {
                    bluePixels++;
                }
            }

            Assert.True(redPixels > 100, "Expected SVG stroke-linejoin=\"bevel\" comparison stroke to render in package preview output.");
            Assert.True(bluePixels > redPixels + 10, $"Expected SVG stroke-linejoin=\"round\" to add visible rounded join coverage beyond bevel joins. Red pixels {redPixels}, blue pixels {bluePixels}.");
        }

        [Fact]
        public void PngRendererProjectsPackageBackedSvgNonScalingStrokePreviewArtwork() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Package SVG Non Scaling Stroke Preview").Size(3, 2);
            string svg = "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"256\" height=\"256\" viewBox=\"0 0 20 20\"><path d=\"M4 6 H16\" fill=\"none\" stroke=\"#dc2626\" stroke-width=\"4\" stroke-linecap=\"butt\"/><path d=\"M4 14 H16\" fill=\"none\" stroke=\"#0070c0\" stroke-width=\"4\" stroke-linecap=\"butt\" vector-effect=\"non-scaling-stroke\"/></svg>";
            AddPackagePreviewShape(page, Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(svg)), "image/svg+xml", ".svg", "../media/non-scaling-stroke.svg");

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            (int RedMin, int RedMax, int RedCount) = FindVerticalColorSpanNear(image, 150, radius: 2, IsRedPixel);
            (int BlueMin, int BlueMax, int BlueCount) = FindVerticalColorSpanNear(image, 150, radius: 2, IsBluePixel);
            Assert.True(RedCount > 0, "Expected regular SVG stroke to render in package preview output.");
            Assert.True(BlueCount > 0, "Expected SVG vector-effect=\"non-scaling-stroke\" stroke to render in package preview output.");
            Assert.True((RedMax - RedMin) > (BlueMax - BlueMin) + 4, $"Expected regular SVG stroke to scale thicker than non-scaling stroke. Red span {RedMin}-{RedMax}, blue span {BlueMin}-{BlueMax}.");
        }

        [Fact]
        public void PngRendererProjectsPackageBackedSvgArcPreviewArtwork() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Package SVG Arc Preview").Size(3, 2);
            string svg = "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 20 20\"><path fill=\"#0070c0\" d=\"M5 2 H15 A3 3 0 0 1 18 5 V15 A3 3 0 0 1 15 18 H5 A3 3 0 0 1 2 15 V5 A3 3 0 0 1 5 2 Z\"/></svg>";
            AddPackagePreviewShape(page, Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(svg)), "image/svg+xml", ".svg", "../media/arc.svg");

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            Assert.True(IsBluePixel(image, 150, 100), "Expected SVG arc preview artwork to render through the native PNG export path.");
            Assert.True(IsWhitePixel(image, 126, 76), "Expected SVG arc preview corners to stay rounded instead of filling as a rectangle.");
        }

        [Fact]
        public void PngRendererProjectsPackageBackedSvgRoundedRectPreviewArtwork() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Package SVG Rounded Rect Preview").Size(3, 2);
            string svg = "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 20 20\"><rect x=\"2\" y=\"2\" width=\"16\" height=\"16\" rx=\"5\" ry=\"5\" fill=\"#0070c0\"/></svg>";
            AddPackagePreviewShape(page, Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(svg)), "image/svg+xml", ".svg", "../media/rounded-rect.svg");

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            Assert.True(IsBluePixel(image, 150, 100), "Expected SVG rounded-rect package preview artwork to render through the native PNG export path.");
            Assert.True(IsWhitePixel(image, 130, 80), "Expected SVG rect rx/ry to preserve the rounded preview corner instead of rendering a sharp rectangle.");
        }

        [Fact]
        public void PngRendererProjectsPackageBackedSvgTransformedPreviewArtwork() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Package SVG Transform Preview").Size(3, 2);
            string svg = "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 20 20\"><path fill=\"#0070c0\" transform=\"rotate(45 10 10)\" d=\"M8 0 H12 V20 H8 Z\"/></svg>";
            AddPackagePreviewShape(page, Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(svg)), "image/svg+xml", ".svg", "../media/transform.svg");

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            Assert.True(HasBluePixelNear(image, 168, 83, radius: 3), "Expected SVG rotate transform to move preview artwork toward the upper-right.");
            Assert.True(IsWhitePixel(image, 150, 78), "Expected SVG rotate transform to remove the unrotated vertical bar from the top center.");
        }

        [Fact]
        public void PngRendererSniffsPackageBackedPreviewArtworkWhenMetadataIsGeneric() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Package Preview Sniff").Size(3, 2);
            AddPackagePreviewShape(page, TrueColorBluePng, "application/octet-stream", ".bin", "../media/blob1.bin");

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            int bluePixels = 0;
            for (int i = 0; i < image.Pixels.Length; i += 4) {
                if (image.Pixels[i] < 60 && image.Pixels[i + 1] < 150 && image.Pixels[i + 2] > 180 && image.Pixels[i + 3] > 200) {
                    bluePixels++;
                }
            }

            Assert.True(bluePixels > 100, "Expected sniffed package preview PNG pixels in the native PNG render.");
        }

        [Fact]
        public void PngRendererNormalizesPackagePreviewContentTypeParameters() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Package Preview Content Type").Size(3, 2);
            AddPackagePreviewShape(page, TrueColorBluePng, "image/png; charset=binary", ".bin", "../media/blob1.bin");

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            int bluePixels = 0;
            for (int i = 0; i < image.Pixels.Length; i += 4) {
                if (image.Pixels[i] < 60 && image.Pixels[i + 1] < 150 && image.Pixels[i + 2] > 180 && image.Pixels[i + 3] > 200) {
                    bluePixels++;
                }
            }

            Assert.True(bluePixels > 100, "Expected parameterized package preview PNG content type to render natively.");
        }

        [Fact]
        public void PngRendererDoesNotUseUnmatchedPackagePreviewRelationships() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Package Preview Stale Metadata").Size(3, 2);
            VisioShape shape = AddPackagePreviewShape(page, TrueColorBluePng);
            shape.SetUserCell("OfficeIMO.StencilPreviewImageRelationshipId", "rIdStale", "STR");
            shape.SetUserCell("OfficeIMO.StencilPreviewImageTarget", "../media/stale.png", "STR");

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            int nonWhitePixels = 0;
            for (int i = 0; i < image.Pixels.Length; i += 4) {
                if (image.Pixels[i] < 245 || image.Pixels[i + 1] < 245 || image.Pixels[i + 2] < 245) {
                    nonWhitePixels++;
                }
            }

            Assert.False(IsBluePixel(image, 150, 100), "Expected stale preview metadata to avoid selecting an unrelated image relationship.");
            Assert.True(nonWhitePixels > 50, "Expected native PNG stencil artwork fallback when package preview metadata cannot be resolved.");
        }

        [Fact]
        public void PngRendererPreservesPackageBackedPngPreviewAspectRatio() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Package Preview Aspect").Size(3, 2);
            AddPackagePreviewShape(page, WideTrueColorBluePng);

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            Assert.True(IsBluePixel(image, 150, 100), "Expected the wide package preview to be centered in the native PNG render.");
            Assert.True(IsWhitePixel(image, 150, 78), "Expected the native PNG renderer to letterbox package previews instead of stretching them vertically.");
            Assert.True(IsWhitePixel(image, 150, 122), "Expected the native PNG renderer to letterbox package previews instead of stretching them vertically.");
        }

        [Fact]
        public void PngRendererRotatesPackagePreviewArtworkWithShape() {
            RgbaPng upright = DecodeRgbaPng(RenderPackagePreviewArtwork(0D));
            RgbaPng rotated = DecodeRgbaPng(RenderPackagePreviewArtwork(Math.PI / 4D));

            Assert.Equal(upright.Width, rotated.Width);
            Assert.Equal(upright.Height, rotated.Height);

            int changedPixels = 0;
            for (int i = 0; i < upright.Pixels.Length; i += 4) {
                int delta = Math.Abs(upright.Pixels[i] - rotated.Pixels[i]) +
                    Math.Abs(upright.Pixels[i + 1] - rotated.Pixels[i + 1]) +
                    Math.Abs(upright.Pixels[i + 2] - rotated.Pixels[i + 2]) +
                    Math.Abs(upright.Pixels[i + 3] - rotated.Pixels[i + 3]);
                if (delta > 80) {
                    changedPixels++;
                }
            }

            Assert.True(changedPixels > 100, "Expected rotated package preview artwork to alter the native PNG render.");
        }

        [Fact]
        public void PngRendererHonorsPackageBackedSvgImageOpacity() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Package SVG Image Opacity").Size(3, 2);
            string svg = "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"10\" height=\"10\" viewBox=\"0 0 10 10\">" +
                "<image x=\"0\" y=\"0\" width=\"10\" height=\"10\" opacity=\"0.5\" href=\"data:image/png;base64," + TrueColorBluePng + "\"/>" +
                "</svg>";
            AddPackagePreviewShape(
                page,
                Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(svg)),
                "image/svg+xml",
                ".svg",
                "../media/image1.svg");

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            Assert.False(IsBluePixel(image, 150, 100), "Expected SVG image opacity to avoid rendering the embedded preview as fully opaque blue.");
            Assert.True(IsPaleBluePixel(image, 150, 100), "Expected package-backed SVG image opacity to blend the embedded preview with the page background.");
        }

        [Fact]
        public void PngRendererUsesPreservedRelativeShapeGeometry() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Preserved Geometry").Size(3, 2);
            AddRelativeTriangleGeometryShape(page);

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            Assert.True(IsWhitePixel(image, 110, 60), "Expected preserved triangle geometry to leave the former rectangle corner untouched.");
            Assert.True(IsRedPixel(image, 150, 100), "Expected preserved triangle geometry to fill the custom imported outline.");
        }

        [Fact]
        public void PngRendererPreservesGeometrySubpathBreaks() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Preserved Subpaths").Size(3, 2);
            AddSubpathBreakGeometryShape(page);

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            Assert.True(IsRedPixel(image, 128, 128), "Expected the first preserved subpath to render.");
            Assert.True(IsRedPixel(image, 173, 82), "Expected the second preserved subpath to render.");
            Assert.True(IsWhitePixel(image, 150, 100), "Expected the break between preserved subpaths to remain unpainted.");
        }

        [Fact]
        public void PngRendererPreservesHolesInPreservedGeometry() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Preserved Geometry Hole").Size(3, 2);
            AddDonutGeometryShape(page);

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            Assert.True(IsRedPixel(image, 120, 100), "Expected the outer preserved contour to fill.");
            Assert.True(IsWhitePixel(image, 150, 100), "Expected the inner preserved contour to cut a hole instead of filling independently.");
        }

        [Fact]
        public void PngRendererLeavesNoFillOpenGeometryUnclosed() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Preserved Open Geometry").Size(3, 2);
            AddOpenNoFillGeometryShape(page);

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            Assert.True(IsRedPixel(image, 150, 130), "Expected the horizontal open geometry stroke to render.");
            Assert.True(IsRedPixel(image, 190, 100), "Expected the vertical open geometry stroke to render.");
            Assert.True(IsWhitePixel(image, 150, 100), "Expected the missing closing edge of open NoFill geometry to remain unpainted.");
        }

        [Fact]
        public void PngRendererSkipsDeletedPreservedGeometryRows() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Preserved Deleted Rows").Size(3, 2);
            AddDeletedGeometryRowShape(page);

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            Assert.True(IsWhitePixel(image, 190, 60), "Expected deleted geometry rows to avoid resurrecting the removed corner.");
            Assert.True(IsRedPixel(image, 150, 100), "Expected the remaining preserved triangle geometry to fill.");
        }

        [Fact]
        public void PngRendererUsesPreservedMasterShapeGeometry() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Master Geometry").Size(3, 2);
            AddMasterBackedTriangleGeometryShape(page);

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            Assert.True(IsWhitePixel(image, 110, 60), "Expected master preserved triangle geometry to leave the former rectangle corner untouched.");
            Assert.True(IsRedPixel(image, 150, 100), "Expected master preserved triangle geometry to fill the scaled imported outline.");
        }

        [Fact]
        public void PngRendererEvaluatesPreservedGeometryWidthHeightFormulas() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Formula Geometry").Size(3, 2);
            AddFormulaGeometryShape(page);

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            Assert.True(IsWhitePixel(image, 110, 100), "Expected formula geometry to leave the former rectangle side untouched.");
            Assert.True(IsRedPixel(image, 150, 100), "Expected formula geometry to fill the evaluated custom outline.");
        }

        [Fact]
        public void PngRendererEvaluatesPreservedGeometryLocPinFormulas() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("LocPin Formula Geometry").Size(3, 2);
            AddLocPinFormulaGeometryShape(page);

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            Assert.True(IsWhitePixel(image, 92, 100), "Expected LocPin formula geometry to leave the former rectangle side untouched.");
            Assert.True(IsRedPixel(image, 150, 100), "Expected LocPin formula geometry to fill the evaluated custom outline.");
        }

        [Fact]
        public void PngRendererEvaluatesPreservedGeometryShapeTransformFormulas() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Shape Transform Formula Geometry").Size(3, 2);
            AddShapeTransformFormulaGeometryShape(page);

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            Assert.True(IsWhitePixel(image, 110, 60), "Expected shape transform formula geometry to leave the former rectangle corner untouched.");
            Assert.True(IsRedPixel(image, 150, 100), "Expected shape transform formula geometry to fill the evaluated custom outline.");
        }

        [Fact]
        public void PngRendererEvaluatesPreservedGeometryMinMaxFormulas() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Min Max Formula Geometry").Size(3, 2);
            AddMinMaxFormulaGeometryShape(page);

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            Assert.True(IsWhitePixel(image, 100, 100), "Expected MIN/MAX formula geometry to leave the former rectangle side untouched.");
            Assert.True(IsRedPixel(image, 150, 100), "Expected MIN/MAX formula geometry to fill the evaluated custom outline.");
        }

        [Fact]
        public void PngRendererEvaluatesPreservedGeometryScalarMathFormulas() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Scalar Math Formula Geometry").Size(3, 2);
            AddScalarMathFormulaGeometryShape(page);

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            Assert.True(IsWhitePixel(image, 100, 100), "Expected ABS/SQRT formula geometry to leave the former rectangle side untouched.");
            Assert.True(IsRedPixel(image, 150, 100), "Expected ABS/SQRT formula geometry to fill the evaluated custom outline.");
        }

        [Fact]
        public void PngRendererEvaluatesPreservedGeometryTrigonometricFormulas() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Trig Formula Geometry").Size(3, 2);
            AddTrigonometricFormulaGeometryShape(page);

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            Assert.True(IsWhitePixel(image, 100, 100), "Expected SIN/COS formula geometry to leave the former rectangle side untouched.");
            Assert.True(IsRedPixel(image, 150, 100), "Expected SIN/COS formula geometry to fill the evaluated custom outline.");
        }

        [Fact]
        public void PngRendererEvaluatesPreservedGeometryAdvancedMathFormulas() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Advanced Math Formula Geometry").Size(3, 2);
            AddAdvancedMathFormulaGeometryShape(page);

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            Assert.True(IsWhitePixel(image, 100, 100), "Expected advanced math formula geometry to leave the former rectangle side untouched.");
            Assert.True(IsRedPixel(image, 150, 100), "Expected advanced math formula geometry to fill the evaluated custom outline.");
        }

        [Fact]
        public void PngRendererEvaluatesPreservedGeometryPowerOperatorFormulas() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Power Operator Formula Geometry").Size(3, 2);
            AddPowerOperatorFormulaGeometryShape(page);

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            Assert.True(IsWhitePixel(image, 100, 100), "Expected power-operator formula geometry to leave the former rectangle side untouched.");
            Assert.True(IsRedPixel(image, 150, 100), "Expected power-operator formula geometry to fill the evaluated custom outline.");
        }

        [Fact]
        public void PngRendererEvaluatesPreservedGeometryUnitFormulas() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Unit Formula Geometry").Size(3, 2);
            AddUnitFormulaGeometryShape(page);

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            Assert.True(IsWhitePixel(image, 100, 100), "Expected unit formula geometry to leave the former rectangle side untouched.");
            Assert.True(IsRedPixel(image, 150, 100), "Expected unit formula geometry to fill the evaluated custom outline.");
        }

        [Fact]
        public void PngRendererEvaluatesPreservedGeometryGuardedFormulas() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Guarded Formula Geometry").Size(3, 2);
            AddGuardedFormulaGeometryShape(page);

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            Assert.True(IsWhitePixel(image, 110, 100), "Expected guarded formula geometry to leave the former rectangle side untouched.");
            Assert.True(IsRedPixel(image, 150, 100), "Expected guarded formula geometry to fill the evaluated custom outline.");
        }

        [Fact]
        public void PngRendererEvaluatesPreservedGeometryIfFormulas() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("If Formula Geometry").Size(3, 2);
            AddIfFormulaGeometryShape(page);

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            Assert.True(IsWhitePixel(image, 100, 100), "Expected IF formula geometry to leave the former rectangle side untouched.");
            Assert.True(IsRedPixel(image, 150, 100), "Expected IF formula geometry to fill the evaluated custom outline.");
        }

        [Fact]
        public void PngRendererEvaluatesOnlySelectedPreservedGeometryIfBranches() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Lazy If Formula Geometry").Size(3, 2);
            AddLazyIfFormulaGeometryShape(page);

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            Assert.True(IsWhitePixel(image, 100, 100), "Expected branch-selected IF formula geometry to leave the former rectangle side untouched.");
            Assert.True(IsRedPixel(image, 150, 100), "Expected branch-selected IF formula geometry to fill the evaluated custom outline.");
        }

        [Fact]
        public void PngRendererEvaluatesPreservedGeometryLogicalIfFormulas() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Logical If Formula Geometry").Size(3, 2);
            AddLogicalIfFormulaGeometryShape(page);

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            Assert.True(IsWhitePixel(image, 100, 100), "Expected logical IF formula geometry to leave the former rectangle side untouched.");
            Assert.True(IsRedPixel(image, 150, 100), "Expected logical IF formula geometry to fill the evaluated custom outline.");
        }

        [Fact]
        public void PngRendererRespectsPreservedGeometryVisibilityFlags() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Geometry Flags").Size(3, 2);
            AddGeometryFlagShape(page);

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            Assert.True(IsWhitePixel(image, 150, 100), "Expected NoFill preserved geometry to leave the triangle interior untouched.");
            Assert.True(IsRedPixel(image, 150, 150), "Expected visible preserved outline geometry to keep its stroke.");
            Assert.True(IsWhitePixel(image, 110, 60), "Expected NoShow preserved geometry to avoid falling back to the former rectangle fill.");
        }

        [Fact]
        public void PngRendererFlattensPreservedEllipseGeometry() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Preserved Ellipse Geometry").Size(3, 2);
            AddEllipseGeometryShape(page);

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            Assert.True(IsWhitePixel(image, 110, 60), "Expected Ellipse geometry to leave the former rectangle corner untouched.");
            Assert.True(IsRedPixel(image, 150, 100), "Expected Ellipse geometry to fill inside the preserved imported outline.");
        }

        [Fact]
        public void PngRendererDrawsPreservedInfiniteLineGeometryAsOpenPath() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Preserved Infinite Line Geometry").Size(3, 2);
            AddInfiniteLineGeometryShape(page);

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            Assert.True(IsRedPixel(image, 150, 100), "Expected InfiniteLine geometry to stroke the clipped diagonal.");
            Assert.True(IsWhitePixel(image, 150, 60), "Expected InfiniteLine geometry to stay open instead of filling the shape interior.");
        }

        [Fact]
        public void PngRendererExpandsPreservedPolylineToGeometry() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Preserved Polyline Geometry").Size(3, 2);
            AddPolylineGeometryShape(page);

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            Assert.True(IsWhitePixel(image, 110, 60), "Expected PolylineTo diamond geometry to leave the former rectangle corner untouched.");
            Assert.True(IsRedPixel(image, 150, 100), "Expected PolylineTo geometry to fill inside the preserved imported outline.");
        }

        [Fact]
        public void PngRendererEvaluatesMinMaxInsidePreservedPolylineFormula() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Polyline Formula Geometry").Size(3, 2);
            AddPolylineMinMaxFormulaGeometryShape(page);

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            Assert.True(IsWhitePixel(image, 100, 60), "Expected POLYLINE MIN/MAX geometry to leave the former rectangle corner untouched.");
            Assert.True(IsRedPixel(image, 150, 100), "Expected POLYLINE MIN/MAX geometry to fill the evaluated diamond outline.");
        }

        [Fact]
        public void PngRendererEvaluatesPercentageLiteralsInsidePreservedPolylineFormula() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Polyline Percent Formula Geometry").Size(3, 2);
            AddPolylinePercentageFormulaGeometryShape(page);

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            Assert.True(IsWhitePixel(image, 100, 60), "Expected POLYLINE percentage geometry to leave the former rectangle corner untouched.");
            Assert.True(IsRedPixel(image, 150, 100), "Expected POLYLINE percentage geometry to fill the evaluated diamond outline.");
        }

        [Fact]
        public void PngRendererFlattensPreservedArcToGeometry() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Preserved Arc Geometry").Size(3, 2);
            AddArcGeometryShape(page);

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            Assert.True(IsWhitePixel(image, 150, 140), "Expected ArcTo geometry to cut away the lower middle instead of rendering a rectangle.");
            Assert.True(IsRedPixel(image, 150, 100), "Expected ArcTo geometry to fill inside the curved imported outline.");
        }

        [Fact]
        public void PngRendererFlattensPreservedEllipticalArcToGeometry() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Preserved Elliptical Arc Geometry").Size(3, 2);
            AddEllipticalArcGeometryShape(page);

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            Assert.True(IsWhitePixel(image, 110, 60), "Expected EllipticalArcTo geometry to leave the former rectangle corner untouched.");
            Assert.True(IsRedPixel(image, 150, 100), "Expected EllipticalArcTo geometry to fill inside the curved imported outline.");
        }

        [Fact]
        public void PngRendererFlattensPreservedRelativeEllipticalArcToGeometry() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Preserved Relative Elliptical Arc Geometry").Size(3, 2);
            AddRelativeEllipticalArcGeometryShape(page);

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            Assert.True(IsWhitePixel(image, 110, 60), "Expected RelEllipticalArcTo geometry to leave the former rectangle corner untouched.");
            Assert.True(IsRedPixel(image, 150, 100), "Expected RelEllipticalArcTo geometry to fill inside the curved imported outline.");
        }

        [Fact]
        public void PngRendererFlattensPreservedRelativeCubicBezierGeometry() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Preserved Cubic Bezier Geometry").Size(3, 2);
            AddRelativeCubicBezierGeometryShape(page);

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            Assert.True(IsWhitePixel(image, 110, 60), "Expected RelCubBezTo geometry to leave the former rectangle corner untouched.");
            Assert.True(IsRedPixel(image, 150, 100), "Expected RelCubBezTo geometry to fill inside the curved imported outline.");
        }

        [Fact]
        public void PngRendererFlattensPreservedAbsoluteCubicBezierGeometry() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Preserved Absolute Cubic Bezier Geometry").Size(3, 2);
            AddAbsoluteCubicBezierGeometryShape(page);

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            Assert.True(IsWhitePixel(image, 110, 60), "Expected CubBezTo geometry to leave the former rectangle corner untouched.");
            Assert.True(IsRedPixel(image, 150, 100), "Expected CubBezTo geometry to fill inside the curved imported outline.");
        }

        [Fact]
        public void PngRendererFlattensPreservedRelativeQuadraticBezierGeometry() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Preserved Quadratic Bezier Geometry").Size(3, 2);
            AddRelativeQuadraticBezierGeometryShape(page);

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            Assert.True(IsWhitePixel(image, 110, 60), "Expected RelQuadBezTo geometry to leave the former rectangle corner untouched.");
            Assert.True(IsRedPixel(image, 150, 120), "Expected RelQuadBezTo geometry to fill inside the curved imported outline.");
        }

        [Fact]
        public void PngRendererFlattensPreservedAbsoluteQuadraticBezierGeometry() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Preserved Absolute Quadratic Bezier Geometry").Size(3, 2);
            AddAbsoluteQuadraticBezierGeometryShape(page);

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            Assert.True(IsWhitePixel(image, 110, 60), "Expected QuadBezTo geometry to leave the former rectangle corner untouched.");
            Assert.True(IsRedPixel(image, 150, 120), "Expected QuadBezTo geometry to fill inside the curved imported outline.");
        }

        [Fact]
        public void PngRendererFlattensPreservedSplineGeometry() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Preserved Spline Geometry").Size(3, 2);
            AddSplineGeometryShape(page);

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            Assert.True(IsWhitePixel(image, 110, 60), "Expected SplineStart/SplineKnot geometry to leave the former rectangle corner untouched.");
            Assert.True(IsRedPixel(image, 150, 100), "Expected SplineStart/SplineKnot geometry to fill inside the curved imported outline.");
        }

        [Fact]
        public void PngRendererSkipsDeletedPreservedSplineKnotRows() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Deleted Spline Knot Geometry").Size(3, 2);
            AddDeletedSplineKnotGeometryShape(page);

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            Assert.True(IsWhitePixel(image, 120, 60), "Expected the deleted SplineKnot row to leave the former rectangle corner untouched.");
            Assert.True(IsRedPixel(image, 180, 100), "Expected the remaining SplineKnot row to keep the preserved triangle filled.");
        }

        [Fact]
        public void PngRendererFlattensPreservedNurbsGeometry() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Preserved NURBS Geometry").Size(3, 2);
            AddNurbsGeometryShape(page);

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            Assert.True(IsWhitePixel(image, 110, 60), "Expected NURBSTo geometry to leave the former rectangle corner untouched.");
            Assert.True(IsRedPixel(image, 150, 100), "Expected NURBSTo geometry to fill inside the curved imported outline.");
        }

        [Fact]
        public void PngRendererUsesVisioCompactNurbsKnotVector() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Preserved Non Uniform NURBS Geometry").Size(3, 2);
            AddNonUniformNurbsGeometryShape(page);

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            bool sawApexPixel = false;
            for (int y = 48; y <= 54; y++) {
                for (int x = 125; x <= 133; x++) {
                    sawApexPixel |= IsRedPixel(image, x, y);
                }
            }

            Assert.True(IsWhitePixel(image, 110, 60), "Expected NURBSTo geometry to leave the former rectangle corner untouched.");
            Assert.True(sawApexPixel, "Expected the compact NURBS knot span to reach the curve apex at the Visio knot boundary.");
        }

        [Fact]
        public void PngRendererProjectsIndexedPackageBackedPngPreviewArtwork() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Indexed Package Preview").Size(3, 2);
            AddPackagePreviewShape(page, IndexedBluePng);

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            int bluePixels = 0;
            for (int i = 0; i < image.Pixels.Length; i += 4) {
                if (image.Pixels[i] < 60 && image.Pixels[i + 1] < 150 && image.Pixels[i + 2] > 180 && image.Pixels[i + 3] > 200) {
                    bluePixels++;
                }
            }

            Assert.True(bluePixels > 100, "Expected indexed package preview PNG pixels in the native PNG render.");
        }

        [Fact]
        public void PngRendererPreservesIndexedPackagePreviewTransparency() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Indexed Transparent Package Preview").Size(3, 2);
            AddPackagePreviewShape(page, IndexedTransparentBluePng);

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            Assert.True(IsWhitePixel(image, 130, 100), "Expected transparent palette preview pixels to leave the native PNG background untouched.");
            Assert.True(IsBluePixel(image, 170, 100), "Expected opaque palette preview pixels to render after transparent pixels are skipped.");
        }

        [Fact]
        public void PngRendererPreservesTrueColorPackagePreviewTransparency() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("TrueColor Transparent Package Preview").Size(3, 2);
            AddPackagePreviewShape(page, TrueColorTransparentBluePng);

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            Assert.True(IsWhitePixel(image, 130, 100), "Expected truecolor tRNS preview pixels to leave the native PNG background untouched.");
            Assert.True(IsBluePixel(image, 170, 100), "Expected non-transparent truecolor preview pixels to render after tRNS matching.");
        }

        [Fact]
        public void PngRendererProjectsGrayscaleAlphaPackageBackedPngPreviewArtwork() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Grayscale Package Preview").Size(3, 2);
            AddPackagePreviewShape(page, GrayscaleAlphaBlackPng);

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            int darkPixels = 0;
            for (int i = 0; i < image.Pixels.Length; i += 4) {
                if (image.Pixels[i] < 50 && image.Pixels[i + 1] < 50 && image.Pixels[i + 2] < 50 && image.Pixels[i + 3] > 200) {
                    darkPixels++;
                }
            }

            Assert.True(darkPixels > 100, "Expected grayscale-alpha package preview PNG pixels in the native PNG render.");
        }

        [Fact]
        public void PngRendererProjectsPackedGrayscalePackageBackedPngPreviewArtwork() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Packed Grayscale Package Preview").Size(3, 2);
            AddPackagePreviewShape(page, PackedGrayscaleBlackPng);

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            int darkPixels = 0;
            for (int i = 0; i < image.Pixels.Length; i += 4) {
                if (image.Pixels[i] < 50 && image.Pixels[i + 1] < 50 && image.Pixels[i + 2] < 50 && image.Pixels[i + 3] > 200) {
                    darkPixels++;
                }
            }

            Assert.True(darkPixels > 100, "Expected packed grayscale package preview PNG pixels in the native PNG render.");
        }

        [Fact]
        public void PngRendererProjectsSixteenBitPackageBackedPngPreviewArtwork() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Sixteen Bit Package Preview").Size(3, 2);
            AddPackagePreviewShape(page, SixteenBitTrueColorBluePng);

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });

            RgbaPng image = DecodeRgbaPng(png);
            int bluePixels = 0;
            for (int i = 0; i < image.Pixels.Length; i += 4) {
                if (image.Pixels[i] < 60 && image.Pixels[i + 1] < 150 && image.Pixels[i + 2] > 180 && image.Pixels[i + 3] > 200) {
                    bluePixels++;
                }
            }

            Assert.True(bluePixels > 100, "Expected 16-bit package preview PNG pixels in the native PNG render.");
        }

        private static void AssertPngHeader(byte[] bytes, int width, int height) {
            Assert.True(bytes.Length > 33);
            AssertPngSignature(bytes);
            Assert.Equal("IHDR", System.Text.Encoding.ASCII.GetString(bytes, 12, 4));
            Assert.Equal(width, ReadBigEndianInt32(bytes, 16));
            Assert.Equal(height, ReadBigEndianInt32(bytes, 20));
            Assert.Equal(8, bytes[24]);
            Assert.Equal(6, bytes[25]);
        }

        private static void AssertPngSignature(byte[] bytes) {
            byte[] signature = { 137, 80, 78, 71, 13, 10, 26, 10 };
            Assert.True(bytes.Length >= signature.Length);
            for (int i = 0; i < signature.Length; i++) {
                Assert.Equal(signature[i], bytes[i]);
            }
        }

        private static int ReadBigEndianInt32(byte[] bytes, int offset) =>
            (bytes[offset] << 24) | (bytes[offset + 1] << 16) | (bytes[offset + 2] << 8) | bytes[offset + 3];

        private static RgbaPng DecodeRgbaPng(byte[] bytes) {
            AssertPngSignature(bytes);
            int width = 0;
            int height = 0;
            using MemoryStream idat = new();
            int offset = 8;
            while (offset < bytes.Length) {
                int length = ReadBigEndianInt32(bytes, offset);
                string type = System.Text.Encoding.ASCII.GetString(bytes, offset + 4, 4);
                int dataOffset = offset + 8;
                if (type == "IHDR") {
                    width = ReadBigEndianInt32(bytes, dataOffset);
                    height = ReadBigEndianInt32(bytes, dataOffset + 4);
                    Assert.Equal(8, bytes[dataOffset + 8]);
                    Assert.Equal(6, bytes[dataOffset + 9]);
                } else if (type == "IDAT") {
                    idat.Write(bytes, dataOffset, length);
                } else if (type == "IEND") {
                    break;
                }

                offset = dataOffset + length + 4;
            }

            byte[] compressed = idat.ToArray();
            using MemoryStream source = new(compressed, 2, compressed.Length - 6);
            using DeflateStream deflate = new(source, CompressionMode.Decompress);
            using MemoryStream inflated = new();
            deflate.CopyTo(inflated);
            byte[] scanlines = inflated.ToArray();
            byte[] rgba = new byte[width * height * 4];
            int sourceOffset = 0;
            int targetOffset = 0;
            for (int y = 0; y < height; y++) {
                Assert.Equal(0, scanlines[sourceOffset++]);
                Buffer.BlockCopy(scanlines, sourceOffset, rgba, targetOffset, width * 4);
                sourceOffset += width * 4;
                targetOffset += width * 4;
            }

            return new RgbaPng(width, height, rgba);
        }

        private static byte[] RenderStyledItalicText(bool italic) {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Text Italic").Size(3, 2);
            VisioShape shape = page.AddRectangle(1.5, 1, 1.8, 0.8, "OfficeIMO");
            shape.FillPattern = 0;
            shape.LinePattern = 0;
            shape.TextStyle = new VisioTextStyle {
                Size = 24,
                TextWidth = 1.6,
                TextHeight = 0.55,
                Italic = italic,
                Color = OfficeColor.FromRgb(22, 101, 52)
            };

            return page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });
        }

        private static byte[] RenderStyledRotatedText(double textAngle) {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Text Rotation").Size(3, 2);
            VisioShape shape = page.AddRectangle(1.5, 1, 1.8, 0.8, "OfficeIMO");
            shape.FillPattern = 0;
            shape.LinePattern = 0;
            shape.TextStyle = new VisioTextStyle {
                Size = 24,
                TextWidth = 1.6,
                TextHeight = 0.55,
                TextAngle = textAngle,
                Color = OfficeColor.FromRgb(22, 101, 52)
            };

            return page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });
        }

        private static byte[] RenderStyledRotatedTextBackground(double textAngle) {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Text Background Rotation").Size(3, 2);
            VisioShape shape = page.AddRectangle(1.5, 1, 1.8, 0.8, "OfficeIMO");
            shape.FillPattern = 0;
            shape.LinePattern = 0;
            shape.TextStyle = new VisioTextStyle {
                Size = 24,
                TextWidth = 1.6,
                TextHeight = 0.55,
                TextAngle = textAngle,
                BackgroundColor = OfficeColor.FromRgb(220, 38, 38),
                BackgroundTransparency = 0,
                Color = OfficeColor.Transparent
            };

            return page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });
        }

        private static byte[] RenderEllipseShape(double angle) {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Ellipse Rotation").Size(3, 2);
            VisioShape shape = page.AddEllipse(1.5, 1, 1.6, 0.55, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LinePattern = 0;
            shape.Angle = angle;

            return page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });
        }

        private static byte[] RenderPackagePreviewArtwork(double angle) {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Package Preview Rotation").Size(3, 2);
            VisioShape shape = AddPackagePreviewShape(page, WideTrueColorBluePng);
            shape.Angle = angle;

            return page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });
        }

        private static byte[] RenderStencilArtwork(double angle) =>
            RenderStencilArtwork(angle, "event.bus");

        private static byte[] RenderStencilArtwork(double angle, string stencilId) {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Stencil Artwork Rotation").Size(3, 2);
            VisioShape shape = page.AddRectangle(1.5, 1, 1.2, 0.8, string.Empty);
            shape.FillPattern = 0;
            shape.LinePattern = 0;
            shape.Angle = angle;
            shape.SetUserCell("OfficeIMO.StencilId", stencilId, "STR");

            return page.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = OfficeColor.White,
                Supersampling = 1
            });
        }

        private static bool IsBluePixel(RgbaPng image, int x, int y) {
            int offset = ((y * image.Width) + x) * 4;
            return image.Pixels[offset] < 60 &&
                   image.Pixels[offset + 1] < 150 &&
                   image.Pixels[offset + 2] > 180 &&
                   image.Pixels[offset + 3] > 200;
        }

        private static bool HasBluePixelNear(RgbaPng image, int x, int y, int radius) {
            int minX = Math.Max(0, x - radius);
            int maxX = Math.Min(image.Width - 1, x + radius);
            int minY = Math.Max(0, y - radius);
            int maxY = Math.Min(image.Height - 1, y + radius);
            for (int py = minY; py <= maxY; py++) {
                for (int px = minX; px <= maxX; px++) {
                    if (IsBluePixel(image, px, py)) {
                        return true;
                    }
                }
            }

            return false;
        }

        private static (int MinX, int MaxX, int Count) FindHorizontalColorSpanNear(RgbaPng image, int y, int radius, Func<RgbaPng, int, int, bool> predicate) {
            int minX = int.MaxValue;
            int maxX = int.MinValue;
            int count = 0;
            int minY = Math.Max(0, y - radius);
            int maxY = Math.Min(image.Height - 1, y + radius);
            for (int py = minY; py <= maxY; py++) {
                for (int px = 0; px < image.Width; px++) {
                    if (predicate(image, px, py)) {
                        minX = Math.Min(minX, px);
                        maxX = Math.Max(maxX, px);
                        count++;
                    }
                }
            }

            return (minX, maxX, count);
        }

        private static (int MinY, int MaxY, int Count) FindVerticalColorSpanNear(RgbaPng image, int x, int radius, Func<RgbaPng, int, int, bool> predicate) {
            int minY = int.MaxValue;
            int maxY = int.MinValue;
            int count = 0;
            int minX = Math.Max(0, x - radius);
            int maxX = Math.Min(image.Width - 1, x + radius);
            for (int px = minX; px <= maxX; px++) {
                for (int py = 0; py < image.Height; py++) {
                    if (predicate(image, px, py)) {
                        minY = Math.Min(minY, py);
                        maxY = Math.Max(maxY, py);
                        count++;
                    }
                }
            }

            return (minY, maxY, count);
        }

        private static bool IsPaleBluePixel(RgbaPng image, int x, int y) {
            int offset = ((y * image.Width) + x) * 4;
            return image.Pixels[offset] > 90 &&
                   image.Pixels[offset] < 180 &&
                   image.Pixels[offset + 1] > 130 &&
                   image.Pixels[offset + 1] < 220 &&
                   image.Pixels[offset + 2] > 220 &&
                   image.Pixels[offset + 3] > 200;
        }

        private static bool IsWhitePixel(RgbaPng image, int x, int y) {
            int offset = ((y * image.Width) + x) * 4;
            return image.Pixels[offset] > 245 &&
                   image.Pixels[offset + 1] > 245 &&
                   image.Pixels[offset + 2] > 245 &&
                   image.Pixels[offset + 3] > 200;
        }

        private static bool IsRedPixel(RgbaPng image, int x, int y) {
            int offset = ((y * image.Width) + x) * 4;
            return image.Pixels[offset] > 180 &&
                   image.Pixels[offset + 1] < 80 &&
                   image.Pixels[offset + 2] < 80 &&
                   image.Pixels[offset + 3] > 200;
        }

        private static bool HasRedPixelNear(RgbaPng image, int x, int y, int radius) {
            int minX = Math.Max(0, x - radius);
            int maxX = Math.Min(image.Width - 1, x + radius);
            int minY = Math.Max(0, y - radius);
            int maxY = Math.Min(image.Height - 1, y + radius);
            for (int py = minY; py <= maxY; py++) {
                for (int px = minX; px <= maxX; px++) {
                    if (IsRedPixel(image, px, py)) {
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool IsGreenPixel(RgbaPng image, int x, int y) {
            int offset = ((y * image.Width) + x) * 4;
            return image.Pixels[offset] < 80 &&
                   image.Pixels[offset + 1] > 80 &&
                   image.Pixels[offset + 1] < 140 &&
                   image.Pixels[offset + 2] < 90 &&
                   image.Pixels[offset + 3] > 200;
        }

        private sealed class RgbaPng {
            internal RgbaPng(int width, int height, byte[] pixels) {
                Width = width;
                Height = height;
                Pixels = pixels;
            }

            internal int Width { get; }

            internal int Height { get; }

            internal byte[] Pixels { get; }
        }

        private const string TrueColorBluePng = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGNgSPj/HwAEIgJfhz+lZwAAAABJRU5ErkJggg==";
        private const string WideTrueColorBluePng = "iVBORw0KGgoAAAANSUhEUgAAAAIAAAABCAYAAAD0In+KAAAADklEQVR4nGNgSPj/H4QBEbsEvYgcJBMAAAAASUVORK5CYII=";
        private const string IndexedBluePng = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABAQMAAAAl21bKAAAABlBMVEX///8AYP8dPVdUAAAAAnRSTlMA/1uRIrUAAAAKSURBVHicY2gAAACCAIF3zXK2AAAAAElFTkSuQmCC";
        private const string IndexedTransparentBluePng = "iVBORw0KGgoAAAANSUhEUgAAAAIAAAABAQMAAADO7O3JAAAABlBMVEX///8AYP8dPVdUAAAAAnRSTlMA/1uRIrUAAAAKSURBVHicY3AAAABCAEEpN/TvAAAAAElFTkSuQmCC";
        private const string TrueColorTransparentBluePng = "iVBORw0KGgoAAAANSUhEUgAAAAIAAAABCAIAAAB7QOjdAAAABnRSTlMA/wAAAP+JwC+QAAAAD0lEQVR4nGP4z/CfIeE/AAu8A14nGkPMAAAAAElFTkSuQmCC";
        private const string GrayscaleAlphaBlackPng = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR4nGNg+A8AAQIBAEK+vGgAAAAASUVORK5CYII=";
        private const string PackedGrayscaleBlackPng = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABAQAAAAA3bvkkAAAACklEQVR4nGNgAAAAAgABSK+kcQAAAABJRU5ErkJggg==";
        private const string SixteenBitTrueColorBluePng = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABEAIAAADA54+dAAAAD0lEQVR4nGNgYEhg+P8fAASEAl/OOhQdAAAAAElFTkSuQmCC";
        private const string TrueColorBlueBmp = "Qk06AAAAAAAAADYAAAAoAAAAAQAAAAEAAAABABgAAAAAAAQAAAATCwAAEwsAAAAAAAAAAAAA/wAAAA==";

        private static VisioShape AddPackagePreviewShape(
            VisioPage page,
            string previewBase64,
            string contentType = "image/png",
            string extension = ".png",
            string target = "../media/image1.png") {
            VisioMaster master = new("package-master", "FancyCloud", new VisioShape("master-shape", 0, 0, 1, 1, string.Empty));
            master.RawMasterRelationships.Add(new VisioAssets.MasterRelationshipContent {
                Id = "rIdImage",
                Type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                Target = target,
                ContentType = contentType,
                Extension = extension,
                Data = Convert.FromBase64String(previewBase64)
            });
            VisioShape shape = page.AddRectangle(1.5, 1, 1.2, 0.8, string.Empty);
            shape.FillPattern = 0;
            shape.LinePattern = 0;
            shape.Master = master;
            shape.NameU = master.NameU;
            shape.SetUserCell("OfficeIMO.StencilId", "package.fancy-cloud", "STR");
            shape.SetUserCell("OfficeIMO.StencilPreviewImageRelationshipId", "rIdImage", "STR");
            shape.SetUserCell("OfficeIMO.StencilPreviewImageTarget", target, "STR");
            shape.SetUserCell("OfficeIMO.StencilPreviewImageContentType", contentType, "STR");
            shape.SetUserCell("OfficeIMO.StencilPreviewImageExtension", extension, "STR");
            return shape;
        }

        private static VisioShape AddRelativeTriangleGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1, 1, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreateRelativeTriangleGeometrySection());
            return shape;
        }

        private static VisioShape AddSubpathBreakGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1, 1, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreateSubpathBreakGeometrySection());
            return shape;
        }

        private static VisioShape AddDonutGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1, 1, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreateDonutGeometrySection());
            return shape;
        }

        private static VisioShape AddOpenNoFillGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1, 1, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineWeight = 0.04D;
            shape.PreservedGeometrySections.Add(CreateOpenNoFillGeometrySection());
            return shape;
        }

        private static VisioShape AddDeletedGeometryRowShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1, 1, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreateDeletedRowGeometrySection());
            return shape;
        }

        private static VisioShape AddArcGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1, 1, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreateArcGeometrySection());
            return shape;
        }

        private static VisioShape AddEllipticalArcGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1, 1, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreateEllipticalArcGeometrySection());
            return shape;
        }

        private static VisioShape AddRelativeEllipticalArcGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1, 1, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreateRelativeEllipticalArcGeometrySection());
            return shape;
        }

        private static VisioShape AddRelativeCubicBezierGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1, 1, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreateRelativeCubicBezierGeometrySection());
            return shape;
        }

        private static VisioShape AddAbsoluteCubicBezierGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1, 1, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreateAbsoluteCubicBezierGeometrySection());
            return shape;
        }

        private static VisioShape AddRelativeQuadraticBezierGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1, 1, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreateRelativeQuadraticBezierGeometrySection());
            return shape;
        }

        private static VisioShape AddAbsoluteQuadraticBezierGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1, 1, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreateAbsoluteQuadraticBezierGeometrySection());
            return shape;
        }

        private static VisioShape AddSplineGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1, 1, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreateSplineGeometrySection());
            return shape;
        }

        private static VisioShape AddDeletedSplineKnotGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1, 1, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreateDeletedSplineKnotGeometrySection());
            return shape;
        }

        private static VisioShape AddNurbsGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1, 1, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreateNurbsGeometrySection());
            return shape;
        }

        private static VisioShape AddNonUniformNurbsGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1, 1, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreateNonUniformNurbsGeometrySection());
            return shape;
        }

        private static VisioShape AddMasterBackedTriangleGeometryShape(VisioPage page) {
            VisioShape masterShape = new("master-shape", 1, 0.5, 2, 1, string.Empty);
            masterShape.PreservedGeometrySections.Add(CreateRelativeTriangleGeometrySection());
            VisioMaster master = new("master-triangle", "PackageTriangle", masterShape);
            VisioShape shape = page.AddShape("master-backed-triangle", master, 1.5, 1, 1, 1, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            return shape;
        }

        private static VisioShape AddFormulaGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1, 1, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreateFormulaTriangleGeometrySection());
            return shape;
        }

        private static VisioShape AddLocPinFormulaGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1.2, 0.8, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreateLocPinFormulaTriangleGeometrySection());
            return shape;
        }

        private static VisioShape AddShapeTransformFormulaGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1, 1, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreateShapeTransformFormulaTriangleGeometrySection());
            return shape;
        }

        private static VisioShape AddMinMaxFormulaGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1.2, 0.8, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreateMinMaxFormulaTriangleGeometrySection());
            return shape;
        }

        private static VisioShape AddScalarMathFormulaGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1.2, 0.8, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreateScalarMathFormulaTriangleGeometrySection());
            return shape;
        }

        private static VisioShape AddTrigonometricFormulaGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1.2, 0.8, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreateTrigonometricFormulaTriangleGeometrySection());
            return shape;
        }

        private static VisioShape AddAdvancedMathFormulaGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1.2, 0.8, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreateAdvancedMathFormulaTriangleGeometrySection());
            return shape;
        }

        private static VisioShape AddPowerOperatorFormulaGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1.2, 0.8, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreatePowerOperatorFormulaTriangleGeometrySection());
            return shape;
        }

        private static VisioShape AddUnitFormulaGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1.2, 0.8, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreateUnitFormulaTriangleGeometrySection());
            return shape;
        }

        private static VisioShape AddGuardedFormulaGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1, 1, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreateGuardedFormulaTriangleGeometrySection());
            return shape;
        }

        private static VisioShape AddIfFormulaGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1.2, 0.8, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreateIfFormulaTriangleGeometrySection());
            return shape;
        }

        private static VisioShape AddLazyIfFormulaGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1.2, 0.8, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreateLazyIfFormulaTriangleGeometrySection());
            return shape;
        }

        private static VisioShape AddLogicalIfFormulaGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1.2, 0.8, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreateLogicalIfFormulaTriangleGeometrySection());
            return shape;
        }

        private static VisioShape AddGeometryFlagShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1, 1, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineWeight = 0.04D;
            shape.PreservedGeometrySections.Add(CreateRelativeTriangleGeometrySection(noFill: true));
            shape.PreservedGeometrySections.Add(CreateFullRectangleGeometrySection(noShow: true));
            return shape;
        }

        private static VisioShape AddEllipseGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1, 1, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreateEllipseGeometrySection());
            return shape;
        }

        private static VisioShape AddInfiniteLineGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1, 1, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(37, 99, 235);
            shape.LineColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineWeight = 0.05D;
            shape.PreservedGeometrySections.Add(CreateInfiniteLineGeometrySection());
            return shape;
        }

        private static VisioShape AddPolylineGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1, 1, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreatePolylineGeometrySection());
            return shape;
        }

        private static VisioShape AddPolylineMinMaxFormulaGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1.2, 0.8, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreatePolylineMinMaxFormulaGeometrySection());
            return shape;
        }

        private static VisioShape AddPolylinePercentageFormulaGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1.2, 0.8, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreatePolylinePercentageFormulaGeometrySection());
            return shape;
        }

        private static XElement CreateRelativeTriangleGeometrySection() {
            return CreateRelativeTriangleGeometrySection(noFill: false);
        }

        private static XElement CreateRelativeTriangleGeometrySection(bool noFill = false, bool noLine = false, bool noShow = false) {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                CreateGeometryRow(ns, noFill, noLine, noShow),
                new XElement(ns + "Row", new XAttribute("T", "RelMoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))),
                new XElement(ns + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))),
                new XElement(ns + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "3"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.5")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "1"))),
                new XElement(ns + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "4"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))));
        }

        private static XElement CreateSubpathBreakGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                CreateGeometryRow(ns, noFill: false, noLine: false, noShow: false),
                new XElement(ns + "Row", new XAttribute("T", "RelMoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.1")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0.1"))),
                new XElement(ns + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.45")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0.1"))),
                new XElement(ns + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "3"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.275")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0.45"))),
                new XElement(ns + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "4"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.1")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0.1"))),
                new XElement(ns + "Row", new XAttribute("T", "RelMoveTo"), new XAttribute("IX", "5"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.55")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0.55"))),
                new XElement(ns + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "6"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.9")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0.55"))),
                new XElement(ns + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "7"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.725")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0.9"))),
                new XElement(ns + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "8"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.55")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0.55"))));
        }

        private static XElement CreateDonutGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                CreateGeometryRow(ns, noFill: false, noLine: false, noShow: false),
                new XElement(ns + "Row", new XAttribute("T", "RelMoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.1")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0.1"))),
                new XElement(ns + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.9")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0.1"))),
                new XElement(ns + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "3"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.9")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0.9"))),
                new XElement(ns + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "4"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.1")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0.9"))),
                new XElement(ns + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "5"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.1")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0.1"))),
                new XElement(ns + "Row", new XAttribute("T", "RelMoveTo"), new XAttribute("IX", "6"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.35")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0.35"))),
                new XElement(ns + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "7"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.65")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0.35"))),
                new XElement(ns + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "8"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.65")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0.65"))),
                new XElement(ns + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "9"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.35")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0.65"))),
                new XElement(ns + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "10"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.35")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0.35"))));
        }

        private static XElement CreateOpenNoFillGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                CreateGeometryRow(ns, noFill: true, noLine: false, noShow: false),
                new XElement(ns + "Row", new XAttribute("T", "RelMoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.1")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0.2"))),
                new XElement(ns + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.9")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0.2"))),
                new XElement(ns + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "3"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.9")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0.8"))));
        }

        private static XElement CreateDeletedRowGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                CreateGeometryRow(ns, noFill: false, noLine: false, noShow: false),
                new XElement(ns + "Row", new XAttribute("T", "RelMoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))),
                new XElement(ns + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))),
                new XElement(ns + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "3"), new XAttribute("Del", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "1"))),
                new XElement(ns + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "4"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.5")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "1"))),
                new XElement(ns + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "5"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))));
        }

        private static XElement CreateFullRectangleGeometrySection(bool noFill = false, bool noLine = false, bool noShow = false) {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "2"),
                CreateGeometryRow(ns, noFill, noLine, noShow),
                new XElement(ns + "Row", new XAttribute("T", "RelMoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))),
                new XElement(ns + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))),
                new XElement(ns + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "3"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "1"))),
                new XElement(ns + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "4"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "1"))),
                new XElement(ns + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "5"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))));
        }

        private static XElement CreateGeometryRow(XNamespace ns, bool noFill, bool noLine, bool noShow) {
            return new XElement(ns + "Row", new XAttribute("T", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Cell", new XAttribute("N", "NoFill"), new XAttribute("V", noFill ? "1" : "0")),
                new XElement(ns + "Cell", new XAttribute("N", "NoLine"), new XAttribute("V", noLine ? "1" : "0")),
                new XElement(ns + "Cell", new XAttribute("N", "NoShow"), new XAttribute("V", noShow ? "1" : "0")));
        }

        private static XElement CreateFormulaTriangleGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Row", new XAttribute("T", "MoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "=Width * 0.25")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "=Height * 0.25"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("F", "Width * 0.75")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "Height / 4"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "3"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "Width / 2")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "Height * 0.75"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "4"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "=Width * 0.25")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "Height * (1 / 4)"))));
        }

        private static XElement CreateLocPinFormulaTriangleGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Row", new XAttribute("T", "MoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("F", "LocPinX - Width * 0.45")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "LocPinY - Height * 0.45"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("F", "LocPinX + Width * 0.45")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "LocPinY - Height * 0.45"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "3"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("F", "LocPinX")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "LocPinY + Height * 0.45"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "4"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("F", "LocPinX - Width * 0.45")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "LocPinY - Height * 0.45"))));
        }

        private static XElement CreateShapeTransformFormulaTriangleGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Row", new XAttribute("T", "MoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("F", "PinX - PinX")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "PinY - PinY"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("F", "PinX - 0.5")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "PinY - PinY"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "3"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("F", "PinX - Width")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "IF(Angle=0, PinY, 0)"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "4"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("F", "PinX - PinX")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "PinY - PinY"))));
        }

        private static XElement CreateMinMaxFormulaTriangleGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Row", new XAttribute("T", "MoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "=MIN(Width, Height) * 0.25")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "=Height * 0.25"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("F", "MAX(Width, Height) * 0.75")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "MIN(Width, Height) * 0.25"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "3"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "Width / 2")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "MIN(Width, Height)"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "4"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "MIN(Width, Height) * 0.25")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "Height * 0.25"))));
        }

        private static XElement CreateScalarMathFormulaTriangleGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Row", new XAttribute("T", "MoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "=ABS(-Height * 0.25)")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "SQRT(Height * Height) * 0.25"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("F", "SQRT(Width * Width) * 0.75")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "ABS(-Height / 4)"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "3"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "Width / 2")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "SQRT(ABS(-Height) * ABS(-Height))"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "4"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "ABS(Height * -0.25)")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "SQRT(Height * Height) / 4"))));
        }

        private static XElement CreateTrigonometricFormulaTriangleGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Row", new XAttribute("T", "MoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "(SIN(PI() - PI()) + COS(0)) * Height * 0.25")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "COS(PI() - PI()) * Height * 0.25"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("F", "COS(0) * Width * 0.75")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "(SIN(0) + COS(0)) * Height * 0.25"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "3"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "Width / 2")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "COS(PI() - PI()) * Height"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "4"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "(SIN(0) + COS(0)) * Height * 0.25")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "COS(0) * Height * 0.25"))));
        }

        private static XElement CreateAdvancedMathFormulaTriangleGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Row", new XAttribute("T", "MoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "(POW(2, 0) + TAN(ATAN2(0, 1)) + TAN(ATAN(0))) * Height * 0.25")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "(RAD(DEG(0)) + 1) * Height * 0.25"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("F", "ROUND(Width * 0.749, 2)")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "INT(Height * 0.9) + Height * 0.25"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "3"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "POW(Width, 1) / 2")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "ROUND(Height * 0.99, 1)"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "4"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "POW(Height, 1) * 0.25")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "(INT(1.9) + TAN(ATAN2(0, 1))) * Height * 0.25"))));
        }

        private static XElement CreateUnitFormulaTriangleGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Row", new XAttribute("T", "MoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "TAN(45 deg) * 0.2 in")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "COS(0 rad) * 0.2 in"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("F", "22.86 mm")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "2.54 cm / 5"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "3"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "Width / (1 in + 1)")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "Height * (1 ft / 12)"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "4"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.2 in")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "2.54 cm / 5"))));
        }

        private static XElement CreatePowerOperatorFormulaTriangleGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Row", new XAttribute("T", "MoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "(Height ^ 2) / Height * 25%")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "(Height ^ 2) / (Height * 4)"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("F", "Width * 75%")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "Height ^ 0 * Height / 4"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "3"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "Width / 2")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "Height ^ 1"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "4"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "Height ^ 2 / Height * 25%")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "Height ^ 0 * Height / 4"))));
        }

        private static XElement CreateGuardedFormulaTriangleGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Row", new XAttribute("T", "MoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "=GUARD(Width * 0.25)")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "GUARD(Height * 0.25)"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("F", "GUARD(Width * 0.75)")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "GUARD(Height / 4)"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "3"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "GUARD(Width / 2)")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "GUARD(Height * 0.75)"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "4"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "GUARD(Width * 0.25)")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "GUARD(Height * (1 / 4))"))));
        }

        private static XElement CreateIfFormulaTriangleGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Row", new XAttribute("T", "MoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "IF(Width > Height, Height * 0.25, Width * 0.25)")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "IF(Width < Height, Width * 0.25, Height * 0.25)"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("F", "IF(Width >= Height, Width * 0.75, Height * 0.75)")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "IF(FALSE, Width, Height * 0.25)"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "3"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "IF(Width = Height, Width, Width / 2)")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "IF(TRUE, Height, Width)"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "4"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "IF(Width <> Height, Height * 0.25, Width * 0.25)")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "IF(Width != Height, Height * 0.25, Width * 0.25)"))));
        }

        private static XElement CreateLazyIfFormulaTriangleGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Row", new XAttribute("T", "MoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "IF(FALSE, Width / 0, Height * 0.25)")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "IF(TRUE, Height * 0.25, Width / 0)"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("F", "IF(TRUE, Width * 0.75, Height / 0)")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "IF(FALSE, Width / 0, Height * 0.25)"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "3"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "IF(Height, Width / 2, Width / 0)")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "IF(Width > Height, IF(TRUE, Height, Width / 0), Width / 0)"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "4"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "IF(FALSE, Width / 0, Height * 0.25)")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "IF(TRUE, MIN(Height, Width) * 0.25, Width / 0)"))));
        }

        private static XElement CreateLogicalIfFormulaTriangleGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Row", new XAttribute("T", "MoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "IF(AND(Width > Height, NOT(FALSE)), Height * 0.25, Width / 0)")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "IF(OR(FALSE, Height > 0), Height * 0.25, Width / 0)"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("F", "IF(AND(TRUE, Width >= Height, NOT(Height > Width)), Width * 0.75, Width / 0)")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "IF(OR(Width < Height, NOT(FALSE)), Height * 0.25, Width / 0)"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "3"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "IF(NOT(Width = Height), Width / 2, Width / 0)")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "IF(AND(OR(FALSE, TRUE), Width > Height), Height, Width / 0)"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "4"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "IF(OR(FALSE, Width <> Height), Height * 0.25, Width / 0)")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "IF(AND(TRUE, NOT(Width < Height)), Height * 0.25, Width / 0)"))));
        }

        private static XElement CreateEllipseGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Row", new XAttribute("T", "Ellipse"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.5")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0.5")),
                    new XElement(ns + "Cell", new XAttribute("N", "A"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "B"), new XAttribute("V", "0.5")),
                    new XElement(ns + "Cell", new XAttribute("N", "C"), new XAttribute("V", "0.5")),
                    new XElement(ns + "Cell", new XAttribute("N", "D"), new XAttribute("V", "1"))));
        }

        private static XElement CreateInfiniteLineGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Row", new XAttribute("T", "InfiniteLine"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "A"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "B"), new XAttribute("V", "1"))));
        }

        private static XElement CreatePolylineGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Row", new XAttribute("T", "MoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.5")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "1"))),
                new XElement(ns + "Row", new XAttribute("T", "PolylineTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.5")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "A"), new XAttribute("F", "POLYLINE(0,0,0.5,1,1,0.5,0.5,0,0,0.5,0.5,1)"))));
        }

        private static XElement CreatePolylineMinMaxFormulaGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Row", new XAttribute("T", "MoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "Width / 2")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "MIN(Width, Height)"))),
                new XElement(ns + "Row", new XAttribute("T", "PolylineTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "Width / 2")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "MIN(Width, Height)")),
                    new XElement(ns + "Cell", new XAttribute("N", "A"), new XAttribute("F", "POLYLINE(1,1,Width/2,MIN(Width,Height),MAX(Width,Height)*0.75,MIN(Width,Height)/2,Width/2,0,MIN(Width,Height)*0.25,MIN(Width,Height)/2,Width/2,MIN(Width,Height))"))));
        }

        private static XElement CreatePolylinePercentageFormulaGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Row", new XAttribute("T", "MoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "50% * Width")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "100% * Height"))),
                new XElement(ns + "Row", new XAttribute("T", "PolylineTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "50% * Width")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "100% * Height")),
                    new XElement(ns + "Cell", new XAttribute("N", "A"), new XAttribute("F", "POLYLINE(1,1,50%*Width,100%*Height,75%*Width,50%*Height,50%*Width,0%*Height,25%*Width,50%*Height,50%*Width,100%*Height)"))));
        }

        private static XElement CreateArcGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Row", new XAttribute("T", "MoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))),
                new XElement(ns + "Row", new XAttribute("T", "ArcTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "A"), new XAttribute("V", "0.25"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "3"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "1"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "4"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "1"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "5"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))));
        }

        private static XElement CreateEllipticalArcGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Row", new XAttribute("T", "MoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))),
                new XElement(ns + "Row", new XAttribute("T", "EllipticalArcTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "A"), new XAttribute("V", "0.5")),
                    new XElement(ns + "Cell", new XAttribute("N", "B"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "C"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "D"), new XAttribute("V", "1"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "3"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))));
        }

        private static XElement CreateRelativeEllipticalArcGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Row", new XAttribute("T", "MoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))),
                new XElement(ns + "Row", new XAttribute("T", "RelEllipticalArcTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "A"), new XAttribute("V", "0.5")),
                    new XElement(ns + "Cell", new XAttribute("N", "B"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "C"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "D"), new XAttribute("V", "1"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "3"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))));
        }

        private static XElement CreateRelativeCubicBezierGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Row", new XAttribute("T", "MoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))),
                new XElement(ns + "Row", new XAttribute("T", "RelCubBezTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "A"), new XAttribute("V", "0.25")),
                    new XElement(ns + "Cell", new XAttribute("N", "B"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "C"), new XAttribute("V", "0.75")),
                    new XElement(ns + "Cell", new XAttribute("N", "D"), new XAttribute("V", "1"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "3"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))));
        }

        private static XElement CreateAbsoluteCubicBezierGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Row", new XAttribute("T", "MoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))),
                new XElement(ns + "Row", new XAttribute("T", "CubBezTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "A"), new XAttribute("V", "0.25")),
                    new XElement(ns + "Cell", new XAttribute("N", "B"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "C"), new XAttribute("V", "0.75")),
                    new XElement(ns + "Cell", new XAttribute("N", "D"), new XAttribute("V", "1"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "3"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))));
        }

        private static XElement CreateRelativeQuadraticBezierGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Row", new XAttribute("T", "MoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))),
                new XElement(ns + "Row", new XAttribute("T", "RelQuadBezTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "A"), new XAttribute("V", "0.5")),
                    new XElement(ns + "Cell", new XAttribute("N", "B"), new XAttribute("V", "1"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "3"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))));
        }

        private static XElement CreateAbsoluteQuadraticBezierGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Row", new XAttribute("T", "MoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))),
                new XElement(ns + "Row", new XAttribute("T", "QuadBezTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "A"), new XAttribute("V", "0.5")),
                    new XElement(ns + "Cell", new XAttribute("N", "B"), new XAttribute("V", "1"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "3"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))));
        }

        private static XElement CreateSplineGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Row", new XAttribute("T", "MoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))),
                new XElement(ns + "Row", new XAttribute("T", "SplineStart"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.25")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "A"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "B"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "C"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "D"), new XAttribute("V", "3"))),
                new XElement(ns + "Row", new XAttribute("T", "SplineKnot"), new XAttribute("IX", "3"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.75")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "A"), new XAttribute("V", "0.5"))),
                new XElement(ns + "Row", new XAttribute("T", "SplineKnot"), new XAttribute("IX", "4"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "A"), new XAttribute("V", "1"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "5"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))));
        }

        private static XElement CreateDeletedSplineKnotGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Row", new XAttribute("T", "MoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))),
                new XElement(ns + "Row", new XAttribute("T", "SplineStart"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "D"), new XAttribute("V", "1"))),
                new XElement(ns + "Row", new XAttribute("T", "SplineKnot"), new XAttribute("IX", "3"), new XAttribute("Del", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "A"), new XAttribute("V", "0.5"))),
                new XElement(ns + "Row", new XAttribute("T", "SplineKnot"), new XAttribute("IX", "4"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "A"), new XAttribute("V", "1"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "5"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))));
        }

        private static XElement CreateNurbsGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Row", new XAttribute("T", "MoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))),
                new XElement(ns + "Row", new XAttribute("T", "NURBSTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "A"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "B"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "C"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "D"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "E"), new XAttribute("F", "NURBS(1,3,0,0,0.25,1,0,1,0.75,1,0,1)"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "3"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))));
        }

        private static XElement CreateNonUniformNurbsGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Row", new XAttribute("T", "MoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))),
                new XElement(ns + "Row", new XAttribute("T", "NURBSTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "A"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "B"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "C"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "D"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "E"), new XAttribute("F", "NURBS(2,2,0,0,0.25,1,0,1,0.5,1,0,1,0.75,0,0.15,1)"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "3"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))));
        }
    }
}
