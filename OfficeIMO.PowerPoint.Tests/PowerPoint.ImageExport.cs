using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests {
    public partial class PowerPointImageExportTests {
        [Fact]
        public void PowerPointPresentation_BatchEncodingConsumesOneSharedByteBudget() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.AddSlide().AddTextBox("First slide");
            presentation.AddSlide().AddTextBox("Second slide");
            var baselineOptions = new PowerPointPresentationImageExportOptions {
                MaximumDegreeOfParallelism = 1
            };
            IReadOnlyList<OfficeImageExportResult> baseline = presentation.ExportImages(
                OfficeImageExportFormat.Png,
                baselineOptions);
            long completeBatchBytes = baseline.Sum(image => image.Bytes.LongLength);
            long maximumBytes = baseline[0].Bytes.LongLength + 1L;

            OfficeImageExportBatchLimitException exception = Assert.Throws<OfficeImageExportBatchLimitException>(() =>
                presentation.ExportImages(
                    OfficeImageExportFormat.Png,
                    new PowerPointPresentationImageExportOptions {
                        MaximumDegreeOfParallelism = 1,
                        MaximumTotalEncodedBytes = maximumBytes
                    }));

            Assert.Equal(nameof(OfficeImageExportOptions.MaximumTotalEncodedBytes), exception.LimitName);
            Assert.Equal(maximumBytes, exception.Maximum);
            Assert.True(exception.Actual < completeBatchBytes);
        }

        [Fact]
        public void PowerPointSlide_DirectImageExportEnforcesRenderTimeout() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            PowerPointSlide slide = presentation.AddSlide();
            var options = new PowerPointImageExportOptions { RenderTimeout = TimeSpan.FromTicks(1) };

            Assert.Throws<OfficeImageExportTimeoutException>(() =>
                slide.ExportImage(OfficeImageExportFormat.Svg, options));
        }

        [Fact]
        public void PowerPointSlide_ExportsSolidBackgroundToPngAndSvgThroughSharedDrawing() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(240, 160);
            PowerPointSlide slide = presentation.AddSlide();
            slide.BackgroundColor = "112233";

            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png, new PowerPointImageExportOptions { Scale = 2D, IncludeSlideContent = false });
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg, new PowerPointImageExportOptions { IncludeSlideContent = false, Scale = 2D });

            Assert.Equal(OfficeImageExportFormat.Png, png.Format);
            Assert.Equal(480, png.Width);
            Assert.Equal(320, png.Height);
            Assert.Equal(480, svg.Width);
            Assert.Equal(320, svg.Height);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(480, image!.Width);
            Assert.Equal(320, image.Height);
            Assert.Equal(OfficeColor.FromRgb(17, 34, 51), image.GetPixel(8, 8));
            AssertNoUnexpectedDiagnostics(png.Diagnostics);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("width=\"480px\"", svgText, StringComparison.Ordinal);
            Assert.Contains("#112233", svgText, StringComparison.OrdinalIgnoreCase);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
        }

        [Fact]
        public void PowerPointSlide_ToImageFluentSavesSvgThroughSharedBuilder() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(240, 160);
            PowerPointSlide slide = presentation.AddSlide();
            slide.BackgroundColor = "112233";
            slide.AddTextBoxPoints("Hidden by fluent content toggle", 20, 20, 140, 24);

            using var output = new MemoryStream();
            OfficeImageExportResult result = slide.ToImage()
                .As(OfficeImageExportFormat.Svg)
                .WithoutContent()
                .Save(output);

            Assert.Equal(OfficeImageExportFormat.Svg, result.Format);
            Assert.Equal(result.Bytes.Length, output.Length);
            string svgText = Encoding.UTF8.GetString(output.ToArray());
            Assert.Contains("#112233", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Hidden by fluent content toggle", svgText, StringComparison.Ordinal);
            AssertNoUnexpectedDiagnostics(result.Diagnostics);
        }

        [Fact]
        public void PowerPointSlide_ToImageUsesConfiguredFormatForBytes() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(160, 90);
            PowerPointSlide slide = presentation.AddSlide();
            slide.BackgroundColor = "112233";

            byte[] png = slide.ToImage()
                .WithoutContent()
                .AsPng()
                .ToBytes();
            string svg = Encoding.UTF8.GetString(slide.ToImage()
                .WithoutContent()
                .AsSvg()
                .ToBytes());

            Assert.Equal(new byte[] { 0x89, 0x50, 0x4E, 0x47 }, png.Take(4).ToArray());
            Assert.Contains("<svg", svg, StringComparison.Ordinal);
            Assert.Contains("#112233", svg, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void PowerPointSlide_FitWithinAndTypedContentOptionsApplyToRasterAndSvg() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(240, 160);
            PowerPointSlide slide = presentation.AddSlide();
            slide.BackgroundColor = "112233";

            OfficeImageExportResult png = slide.ToImage()
                .WithScale(4D)
                .FitWithin(300, 300)
                .IncludePictures(false)
                .IncludeCharts(false)
                .AsPng()
                .Export();
            OfficeImageExportResult svg = slide.ToImage()
                .WithScale(4D)
                .FitWithin(300, 300)
                .ConfigureOptions(options => options.IncludeHiddenShapes = false)
                .AsSvg()
                .Export();

            Assert.Equal((300, 200), (png.Width, png.Height));
            Assert.Equal((png.Width, png.Height), (svg.Width, svg.Height));
        }

        [Theory]
        [InlineData(0)]
        [InlineData(-1)]
        public void PowerPointSlide_GroupDepthFluentAcceptsZeroAndUnlimitedDepth(int maximumDepth) {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            PowerPointSlide slide = presentation.AddSlide();

            OfficeImageExportResult result = slide.ToImage()
                .WithMaximumGroupDepth(maximumDepth)
                .AsSvg()
                .Export();

            Assert.Equal(OfficeImageExportFormat.Svg, result.Format);
        }

        [Fact]
        public void PowerPointPresentation_ToImagesFluentSavesVisibleSlidesAndCanIncludeHiddenSlides() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(180, 100);

            PowerPointSlide visible = presentation.AddSlide();
            visible.BackgroundColor = "112233";

            PowerPointSlide hidden = presentation.AddSlide();
            hidden.Hidden = true;
            hidden.BackgroundColor = "445566";

            string folder = Path.Combine(Path.GetTempPath(), "officeimo-ppt-images-" + Guid.NewGuid().ToString("N"));
            try {
                IReadOnlyList<OfficeImageExportResult> visibleResults = presentation.ToImages()
                    .AsSvg()
                    .Save(folder);

                Assert.Single(visibleResults);
                Assert.Equal("Slide 1", visibleResults[0].Name);
                string visiblePath = Path.Combine(folder, "Slide 1.svg");
                Assert.True(File.Exists(visiblePath));
                string visibleSvg = File.ReadAllText(visiblePath);
                Assert.Contains("#112233", visibleSvg, StringComparison.OrdinalIgnoreCase);
                Assert.DoesNotContain("#445566", visibleSvg, StringComparison.OrdinalIgnoreCase);
                Assert.False(File.Exists(Path.Combine(folder, "Slide 2.svg")));

                string allFolder = Path.Combine(folder, "all");
                IReadOnlyList<OfficeImageExportResult> allResults = presentation.ToImages()
                    .IncludeHiddenSlides()
                    .AsSvg()
                    .Save(allFolder);

                Assert.Equal(2, allResults.Count);
                Assert.Equal("Slide 2", allResults[1].Name);
                string hiddenPath = Path.Combine(allFolder, "Slide 2.svg");
                Assert.True(File.Exists(hiddenPath));
                Assert.Contains("#445566", File.ReadAllText(hiddenPath), StringComparison.OrdinalIgnoreCase);
            } finally {
                if (Directory.Exists(folder)) {
                    Directory.Delete(folder, recursive: true);
                }
            }
        }

        [Fact]
        public void PowerPointPresentation_ToImagesFluentSelectsSlidesAndRangesThroughSharedBatchBuilder() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(180, 100);

            PowerPointSlide first = presentation.AddSlide();
            first.BackgroundColor = "112233";
            first.AddTextBoxPoints("First selected", 20, 20, 120, 24);

            PowerPointSlide hidden = presentation.AddSlide();
            hidden.Hidden = true;
            hidden.BackgroundColor = "445566";
            hidden.AddTextBoxPoints("Hidden selected", 20, 20, 120, 24);

            PowerPointSlide third = presentation.AddSlide();
            third.BackgroundColor = "778899";
            third.AddTextBoxPoints("Third selected", 20, 20, 120, 24);

            string folder = Path.Combine(Path.GetTempPath(), "officeimo-ppt-filtered-images-" + Guid.NewGuid().ToString("N"));
            try {
                IReadOnlyList<OfficeImageExportResult> selectedResults = presentation.ToImages()
                    .ForSlides(3, 1, 3)
                    .AsSvg()
                    .Save(folder);

                Assert.Equal(2, selectedResults.Count);
                Assert.Equal("Slide 1", selectedResults[0].Name);
                Assert.Equal("Slide 3", selectedResults[1].Name);
                Assert.True(File.Exists(Path.Combine(folder, "Slide 1.svg")));
                Assert.False(File.Exists(Path.Combine(folder, "Slide 2.svg")));
                Assert.True(File.Exists(Path.Combine(folder, "Slide 3.svg")));
                Assert.Contains("#112233", File.ReadAllText(Path.Combine(folder, "Slide 1.svg")), StringComparison.OrdinalIgnoreCase);
                Assert.Contains("#778899", File.ReadAllText(Path.Combine(folder, "Slide 3.svg")), StringComparison.OrdinalIgnoreCase);

                string rangeFolder = Path.Combine(folder, "range");
                IReadOnlyList<OfficeImageExportResult> rangeResults = presentation.ToImages()
                    .ForSlideRange(2, 3)
                    .IncludeHiddenSlides()
                    .AsSvg()
                    .Save(rangeFolder);

                Assert.Equal(2, rangeResults.Count);
                Assert.Equal("Slide 2", rangeResults[0].Name);
                Assert.Equal("Slide 3", rangeResults[1].Name);
                Assert.Contains("#445566", File.ReadAllText(Path.Combine(rangeFolder, "Slide 2.svg")), StringComparison.OrdinalIgnoreCase);
            } finally {
                if (Directory.Exists(folder)) {
                    Directory.Delete(folder, recursive: true);
                }
            }
        }

        [Fact]
        public void PowerPointSlide_ExportsGradientBackgroundThroughDrawingGradient() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(120, 80);
            PowerPointSlide slide = presentation.AddSlide();
            slide.SetBackgroundGradient("112233", "445566", 45D);

            OfficeImageExportResult result = slide.ExportImage(OfficeImageExportFormat.Svg, new PowerPointImageExportOptions { IncludeSlideContent = false });
            string svgText = Encoding.UTF8.GetString(result.Bytes);

            Assert.Contains("<linearGradient", svgText, StringComparison.Ordinal);
            Assert.Contains("#112233", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#445566", svgText, StringComparison.OrdinalIgnoreCase);
            AssertNoUnexpectedDiagnostics(result.Diagnostics);
        }

        [Fact]
        public void PowerPointSlide_SharedSnapshotStopsAtConfiguredGroupDepth() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape first = slide.AddShapePoints(OfficePresetShapeType.Rectangle, 10, 10, 20, 20);
            PowerPointAutoShape second = slide.AddShapePoints(OfficePresetShapeType.Rectangle, 40, 10, 20, 20);
            PowerPointAutoShape third = slide.AddShapePoints(OfficePresetShapeType.Rectangle, 70, 10, 20, 20);
            PowerPointGroupShape inner = slide.GroupShapes(new PowerPointShape[] { first, second }, "Inner");
            slide.GroupShapes(new PowerPointShape[] { inner, third }, "Outer");

            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot(
                new PowerPointImageExportOptions { MaxGroupShapeDepth = 1 });

            Assert.Contains(snapshot.Diagnostics, diagnostic =>
                diagnostic.Message.Contains(nameof(PowerPointImageExportOptions.MaxGroupShapeDepth), StringComparison.Ordinal));
        }

        [Fact]
        public void PowerPointSlide_RendersRectangleAndTextBoxThroughSharedDrawing() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(160, 100);
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape rectangle = slide.AddRectanglePoints(20, 20, 60, 30);
            rectangle.FillColor = "22AA66";
            rectangle.OutlineColor = "114433";
            rectangle.OutlineWidthPoints = 1D;

            PowerPointTextBox textBox = slide.AddTextBoxPoints("Shared renderer", 24, 54, 150, 36);
            textBox.FontSize = 12;
            textBox.Color = "111111";

            OfficeImageExportResult result = slide.ExportImage(OfficeImageExportFormat.Png);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            AssertNoUnexpectedDiagnostics(result.Diagnostics);
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            Assert.True(OfficePngReader.TryDecode(result.Bytes, out OfficeRasterImage? image));
            Assert.Equal(OfficeColor.FromRgb(34, 170, 102), image!.GetPixel(30, 30));
            Assert.Contains("Shared renderer", GetSnapshotPlainText(snapshot), StringComparison.Ordinal);
        }

        [Fact]
        public void PowerPointSlide_RendersPictureThroughSharedDrawingImageElement() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(160, 100);
            PowerPointSlide slide = presentation.AddSlide();
            byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(2, 2, OfficeColor.CornflowerBlue));
            slide.AddPicture(new MemoryStream(png), OfficeImageFormat.Png, PowerPointUnits.FromPoints(20), PowerPointUnits.FromPoints(20), PowerPointUnits.FromPoints(20), PowerPointUnits.FromPoints(20));

            OfficeImageExportResult pngResult = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svgResult = slide.ExportImage(OfficeImageExportFormat.Svg);

            AssertNoUnexpectedDiagnostics(pngResult.Diagnostics);
            AssertNoUnexpectedDiagnostics(svgResult.Diagnostics);
            Assert.True(OfficePngReader.TryDecode(pngResult.Bytes, out OfficeRasterImage? image));
            Assert.Equal(OfficeColor.CornflowerBlue, image!.GetPixel(25, 25));

            string svgText = Encoding.UTF8.GetString(svgResult.Bytes);
            Assert.Equal(OfficeColor.CornflowerBlue, ReadSvgOutput(svgResult).GetPixel(25, 25));
        }

        [Fact]
        public void PowerPointSlide_BoundsEmbeddedPictureReads() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            PowerPointSlide slide = presentation.AddSlide();
            byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(2, 2, OfficeColor.CornflowerBlue));
            slide.AddPicture(new MemoryStream(png), OfficeImageFormat.Png,
                PowerPointUnits.FromPoints(20), PowerPointUnits.FromPoints(20),
                PowerPointUnits.FromPoints(20), PowerPointUnits.FromPoints(20));

            OfficeImageExportResult result = slide.ExportImage(OfficeImageExportFormat.Png, new PowerPointImageExportOptions {
                MaximumEmbeddedImageBytes = png.Length - 1,
                MaximumTotalEmbeddedImageBytes = png.Length * 2L
            });

            Assert.True(OfficePngReader.TryDecode(result.Bytes, out OfficeRasterImage? image));
            Assert.Equal(0, CountPixelsNear(image!, OfficeColor.CornflowerBlue));
            Assert.Contains(result.Diagnostics, diagnostic =>
                diagnostic.LossKind == OfficeConversionLossKind.Omission);
        }

        [Fact]
        public void PowerPointSlide_BoundsAggregateEmbeddedPictureReads() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            PowerPointSlide slide = presentation.AddSlide();
            byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(2, 2, OfficeColor.CornflowerBlue));
            slide.AddPicture(new MemoryStream(png), OfficeImageFormat.Png,
                PowerPointUnits.FromPoints(20), PowerPointUnits.FromPoints(20),
                PowerPointUnits.FromPoints(20), PowerPointUnits.FromPoints(20));
            slide.AddPicture(new MemoryStream(png), OfficeImageFormat.Png,
                PowerPointUnits.FromPoints(50), PowerPointUnits.FromPoints(20),
                PowerPointUnits.FromPoints(20), PowerPointUnits.FromPoints(20));

            OfficeImageExportResult result = slide.ExportImage(OfficeImageExportFormat.Png, new PowerPointImageExportOptions {
                MaximumEmbeddedImageBytes = png.Length,
                MaximumTotalEmbeddedImageBytes = png.Length
            });

            Assert.True(OfficePngReader.TryDecode(result.Bytes, out OfficeRasterImage? image));
            Assert.Equal(OfficeColor.CornflowerBlue, image!.GetPixel(25, 25));
            Assert.NotEqual(OfficeColor.CornflowerBlue, image.GetPixel(55, 25));
            Assert.Contains(result.Diagnostics, diagnostic => diagnostic.LossKind == OfficeConversionLossKind.Omission);
        }

        [Fact]
        public void PowerPointSlide_ChargesGroupedPicturesOnceAgainstAggregateBudget() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            PowerPointSlide slide = presentation.AddSlide();
            byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(2, 2, OfficeColor.CornflowerBlue));
            PowerPointPicture picture = slide.AddPicture(new MemoryStream(png), OfficeImageFormat.Png,
                PowerPointUnits.FromPoints(20), PowerPointUnits.FromPoints(20),
                PowerPointUnits.FromPoints(20), PowerPointUnits.FromPoints(20));
            PowerPointAutoShape anchor = slide.AddRectanglePoints(50, 20, 12, 12);
            slide.GroupShapes(new PowerPointShape[] { picture, anchor }, "Bounded picture group");

            OfficeImageExportResult result = slide.ExportImage(OfficeImageExportFormat.Png, new PowerPointImageExportOptions {
                MaximumEmbeddedImageBytes = png.Length,
                MaximumTotalEmbeddedImageBytes = png.Length
            });

            Assert.True(OfficePngReader.TryDecode(result.Bytes, out OfficeRasterImage? image));
            Assert.Equal(OfficeColor.CornflowerBlue, image!.GetPixel(25, 25));
            Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.LossKind == OfficeConversionLossKind.Omission);
        }

        [Fact]
        public void PowerPointSlide_BoundsEmbeddedBackgroundImageReads() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            PowerPointSlide slide = presentation.AddSlide();
            byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(2, 2, OfficeColor.CornflowerBlue));
            slide.SetBackgroundImage(new MemoryStream(png), OfficeImageFormat.Png);

            OfficeImageExportResult result = slide.ExportImage(OfficeImageExportFormat.Png, new PowerPointImageExportOptions {
                MaximumEmbeddedImageBytes = png.Length - 1,
                MaximumTotalEmbeddedImageBytes = png.Length * 2L
            });

            Assert.True(OfficePngReader.TryDecode(result.Bytes, out OfficeRasterImage? image));
            Assert.Equal(0, CountPixelsNear(image!, OfficeColor.CornflowerBlue));
            Assert.Contains(result.Diagnostics, diagnostic => diagnostic.LossKind == OfficeConversionLossKind.Omission);
        }

        [Fact]
        public void PowerPointSlide_RasterizesExistingSvgPictureWithoutAPlaceholder() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(160, 100);
            PowerPointSlide slide = presentation.AddSlide();
            byte[] svg = Encoding.UTF8.GetBytes(
                "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 10 10\">" +
                "<rect x=\"0\" y=\"0\" width=\"10\" height=\"10\" fill=\"#D7263D\"/>" +
                "</svg>");
            slide.AddPicture(
                new MemoryStream(svg),
                OfficeImageFormat.Svg,
                PowerPointUnits.FromPoints(20),
                PowerPointUnits.FromPoints(20),
                PowerPointUnits.FromPoints(40),
                PowerPointUnits.FromPoints(30));

            OfficeImageExportResult result = slide.ExportImage(OfficeImageExportFormat.Png);

            Assert.DoesNotContain(
                result.Diagnostics,
                diagnostic => diagnostic.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodeFallback);
            Assert.True(OfficePngReader.TryDecode(result.Bytes, out OfficeRasterImage? image));
            Assert.Equal(OfficeColor.FromRgb(0xD7, 0x26, 0x3D), image!.GetPixel(40, 35));
        }

        [Fact]
        public void PowerPointSlide_RendersImageBackgroundThroughSharedDrawingImageElement() {
            string imagePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".png");
            try {
                OfficeColor backgroundImageColor = OfficeColor.FromRgb(60, 179, 113);
                byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(4, 4, backgroundImageColor));
                File.WriteAllBytes(imagePath, png);

                using var stream = new MemoryStream();
                using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
                presentation.SlideSize.SetSizePoints(80, 60);
                PowerPointSlide slide = presentation.AddSlide();
                slide.SetBackgroundImage(imagePath);

                OfficeImageExportResult result = slide.ExportImage(OfficeImageExportFormat.Png, new PowerPointImageExportOptions { IncludeSlideContent = false });

                AssertNoUnexpectedDiagnostics(result.Diagnostics);
                Assert.True(OfficePngReader.TryDecode(result.Bytes, out OfficeRasterImage? image));
                Assert.Equal(backgroundImageColor, image!.GetPixel(5, 5));
            } finally {
                if (File.Exists(imagePath)) {
                    File.Delete(imagePath);
                }
            }
        }

        [Fact]
        public void PowerPointSlide_RendersTableCellsThroughSharedDrawing() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(180, 120);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointTable table = slide.AddTablePoints(2, 2, 20, 20, 120, 50);
            table.SetColumnWidthsPoints(45, 75);
            table.SetRowHeightsPoints(20, 30);
            table.Rotation = 5D;

            PowerPointTableCell header = table.GetCell(0, 0);
            header.Text = "Region";
            header.FillColor = "22AA66";
            header.BorderColor = "114433";
            header.Color = "FFFFFF";
            header.Bold = true;
            header.FontSize = 11;
            header.HorizontalAlignment = PowerPointTextAlignment.Center;
            header.VerticalAlignment = PowerPointTextVerticalAlignment.Center;

            PowerPointTableCell merged = table.GetCell(1, 0);
            merged.Text = "Shared table renderer";
            merged.FillColor = "EAF7F0";
            merged.BorderColor = "114433";
            merged.Merge = (1, 2);
            table.GetCell(1, 1).Text = string.Empty;

            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            AssertNoUnexpectedDiagnostics(png.Diagnostics, "table-cell-text-fit-approximation");
            AssertNoUnexpectedDiagnostics(svg.Diagnostics, "table-cell-text-fit-approximation");
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics, "table-cell-text-fit-approximation");
            Assert.Contains("Region", GetSnapshotPlainText(snapshot), StringComparison.Ordinal);
            Assert.Contains("Shared table renderer", GetSnapshotPlainText(snapshot), StringComparison.Ordinal);

            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(png.Width, image!.Width);
            Assert.Equal(png.Height, image.Height);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("#22AA66", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.True(CountPixelsNear(image, OfficeColor.FromRgb(34, 170, 102)) > 20);
        }

        [Fact]
        public void PowerPointSlide_ExportsListMarkersInPdfTextSpans() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(320, 260);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointTextBox bullets = slide.AddTextBoxPoints(string.Empty, 20, 20, 280, 220);
            bullets.SetBullets(new[] { "First bullet", "Second bullet" });
            bullets.AddNumberedList(new[] { "First number", "Second number" }, startAt: 3);
            bullets.AddBullets(new[] { "Nested bullet" }, level: 1);

            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            string text = GetSnapshotPlainText(snapshot);
            Assert.Contains("First bullet", text, StringComparison.Ordinal);
            Assert.Contains("Second bullet", text, StringComparison.Ordinal);
            Assert.Contains("3.", text, StringComparison.Ordinal);
            Assert.Contains("4.", text, StringComparison.Ordinal);
            Assert.Contains("Nested bullet", text, StringComparison.Ordinal);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(320, image!.Width);
            Assert.Equal(260, image.Height);

        }

        [Fact]
        public void PowerPointSlide_ExportsAutoNumberSchemesInPdfTextSpans() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(340, 260);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointTextBox alpha = slide.AddTextBoxPoints(string.Empty, 20, 20, 300, 90);
            alpha.AddNumberedList(
                new[] { "Alpha lower", "Alpha next" },
                PowerPointNumberingScheme.AlphaLowerCharacterPeriod,
                startAt: 2);

            PowerPointTextBox roman = slide.AddTextBoxPoints(string.Empty, 20, 130, 300, 90);
            roman.AddNumberedList(
                new[] { "Roman upper", "Roman next" },
                PowerPointNumberingScheme.RomanUpperCharacterParenR,
                startAt: 4);

            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            string text = GetSnapshotPlainText(snapshot);
            Assert.Contains("b.", text, StringComparison.Ordinal);
            Assert.Contains("c.", text, StringComparison.Ordinal);
            Assert.Contains("IV)", text, StringComparison.Ordinal);
            Assert.Contains("V)", text, StringComparison.Ordinal);
            Assert.Contains("Roman next", text, StringComparison.Ordinal);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(340, image!.Width);
            Assert.Equal(260, image.Height);

        }

        [Fact]
        public void PowerPointSlide_KeepsLayoutOnlyPlaceholderTextInTheTemplate() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(240, 160);

            int layoutIndex = presentation.GetLayoutIndex(PowerPointSlideLayoutType.TitleOnly);
            var bounds = new PowerPointLayoutBox(
                PowerPointUnits.FromPoints(20),
                PowerPointUnits.FromPoints(24),
                PowerPointUnits.FromPoints(180),
                PowerPointUnits.FromPoints(32));
            PowerPointTextBox layoutTitle = presentation.EnsureLayoutPlaceholderTextBox(0, layoutIndex, PowerPointPlaceholderType.Title, bounds: bounds);
            presentation.SetLayoutPlaceholderBounds(0, layoutIndex, PowerPointPlaceholderType.Title, bounds);
            layoutTitle.Text = "Inherited layout title";
            layoutTitle.FontSize = 14;
            layoutTitle.Color = "1F4E79";

            PowerPointSlide slide = presentation.AddSlide(PowerPointSlideLayoutType.TitleOnly);

            OfficeImageExportResult result = slide.ExportImage(OfficeImageExportFormat.Png);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            Assert.Empty(slide.Shapes);
            AssertNoUnexpectedDiagnostics(result.Diagnostics);
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            Assert.DoesNotContain("Inherited layout title", GetSnapshotPlainText(snapshot), StringComparison.Ordinal);
            Assert.True(OfficePngReader.TryDecode(result.Bytes, out OfficeRasterImage? image));
            Assert.Equal(240, image!.Width);
            Assert.Equal(160, image.Height);
        }

        [Fact]
        public void PowerPointSlide_SuppressesInheritedLayoutPlaceholderWhenSlidePlaceholderOverridesIt() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(240, 160);

            int layoutIndex = presentation.GetLayoutIndex(PowerPointSlideLayoutType.TitleOnly);
            var bounds = new PowerPointLayoutBox(
                PowerPointUnits.FromPoints(20),
                PowerPointUnits.FromPoints(24),
                PowerPointUnits.FromPoints(180),
                PowerPointUnits.FromPoints(32));
            PowerPointTextBox layoutTitle = presentation.EnsureLayoutPlaceholderTextBox(0, layoutIndex, PowerPointPlaceholderType.Title, bounds: bounds);
            presentation.SetLayoutPlaceholderBounds(0, layoutIndex, PowerPointPlaceholderType.Title, bounds);
            layoutTitle.Text = "Inherited layout title";
            layoutTitle.Color = "C00000";

            PowerPointSlide slide = presentation.AddSlide(PowerPointSlideLayoutType.TitleOnly);
            PowerPointTextBox slideTitle = slide.AddTitlePoints("Slide local title", 20, 24, 180, 32);
            slideTitle.Color = "1F4E79";

            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            Assert.Single(slide.Shapes);
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            Assert.Contains("Slide local title", GetSnapshotPlainText(snapshot), StringComparison.Ordinal);
            Assert.DoesNotContain("Inherited layout title", GetSnapshotPlainText(snapshot), StringComparison.Ordinal);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(240, image!.Width);
            Assert.Equal(160, image.Height);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Slide local title", GetSnapshotPlainText(snapshot), StringComparison.Ordinal);
            Assert.DoesNotContain("Inherited layout title", GetSnapshotPlainText(snapshot), StringComparison.Ordinal);
        }

        [Fact]
        public void PowerPointImageExportOptionsReuseSharedOfficeImageExportOptions() {
            PowerPointImageExportOptions options = new PowerPointImageExportOptions {
                Scale = 1.25D,
                BackgroundColor = OfficeColor.AliceBlue
            };

            Assert.IsAssignableFrom<OfficeImageExportOptions>(options);
            Assert.Equal(1.25D, options.Scale);
            Assert.Equal(OfficeColor.AliceBlue, options.BackgroundColor);

            PowerPointPresentationImageExportOptions presentationOptions = new PowerPointPresentationImageExportOptions {
                Scale = 1.5D,
                BackgroundColor = OfficeColor.White,
                IncludeHiddenSlides = true,
                SlideNumbers = new[] { 1, 3 }
            };

            Assert.IsAssignableFrom<OfficeImageExportOptions>(presentationOptions);
            Assert.Equal(1.5D, presentationOptions.Scale);
            Assert.Equal(OfficeColor.White, presentationOptions.BackgroundColor);
            Assert.True(presentationOptions.IncludeHiddenSlides);
            Assert.Equal(new[] { 1, 3 }, presentationOptions.SlideNumbers);

            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                presentation.AddSlide().ExportImage(OfficeImageExportFormat.Png, new PowerPointImageExportOptions { Scale = 0D }));
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                presentation.ExportImages(OfficeImageExportFormat.Png, new PowerPointPresentationImageExportOptions { SlideNumbers = new[] { 0 } }));
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                presentation.ExportImages(OfficeImageExportFormat.Png, new PowerPointPresentationImageExportOptions { SlideNumbers = new[] { 2 } }));
            Assert.Throws<ArgumentOutOfRangeException>(() => presentation.ToImages().ForSlides(0));
            Assert.Throws<ArgumentOutOfRangeException>(() => presentation.ToImages().ForSlideRange(2, 1));
        }

        [Fact]
        public void PresentationImageExportPropagatesAllSlideContentFilters() {
            string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "BackgroundImage.png");
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            PowerPointSlide slide = presentation.AddSlide();
            slide.AddPicture(imagePath);
            PowerPointAutoShape shape = slide.AddRectanglePoints(12, 12, 30, 20, "Filtered shape");
            shape.FillColor = "ABCDEF";
            slide.AddTextBoxPoints("FILTERED TEXT", 48, 12, 100, 24);
            PowerPointTable table = slide.AddTablePoints(1, 1, 12, 44, 120, 30);
            table.GetCell(0, 0).Text = "FILTERED TABLE";
            slide.AddChartPoints(OfficeChartKind.ColumnClustered,
                new OfficeChartData(new[] { "FILTERED CATEGORY" }, new[] {
                    new OfficeChartSeries("FILTERED CHART", new[] { 42D })
                }), 150, 44, 120, 80);

            OfficeImageExportResult result = Assert.Single(presentation.ExportImages(
                OfficeImageExportFormat.Svg, new PowerPointPresentationImageExportOptions {
                    IncludeSlideBackground = false,
                    IncludePictures = false,
                    IncludeAutoShapes = false,
                    IncludeTextBoxes = false,
                    IncludeTables = false,
                    IncludeCharts = false
                }));
            string svg = Encoding.UTF8.GetString(result.Bytes);

            Assert.DoesNotContain("<image", svg, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("ABCDEF", svg, StringComparison.OrdinalIgnoreCase);
            // A blank render proves filtering even when SVG contains glyph outlines rather than text.
            OfficeImageExportResult png = Assert.Single(presentation.ExportImages(
                OfficeImageExportFormat.Png, new PowerPointPresentationImageExportOptions {
                    IncludeSlideBackground = false, IncludePictures = false, IncludeAutoShapes = false,
                    IncludeTextBoxes = false, IncludeTables = false, IncludeCharts = false,
                    BackgroundColor = OfficeColor.White
                }));
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(image!.Width * image.Height, CountPixelsNear(image, OfficeColor.White));
        }

        private static int CountPixelsNear(OfficeRasterImage image, OfficeColor expected) {
            int count = 0;
            for (int y = 0; y < image.Height; y++) {
                for (int x = 0; x < image.Width; x++) {
                    OfficeColor actual = image.GetPixel(x, y);
                    if (Math.Abs(actual.R - expected.R) <= 8 &&
                        Math.Abs(actual.G - expected.G) <= 8 &&
                        Math.Abs(actual.B - expected.B) <= 8 &&
                        Math.Abs(actual.A - expected.A) <= 8) {
                        count++;
                    }
                }
            }

            return count;
        }

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

        private static byte[] CreateBmp32(int width, int height, IReadOnlyList<OfficeColor> pixels) {
            int rowStride = width * 4;
            int pixelOffset = 54;
            byte[] bytes = new byte[pixelOffset + (rowStride * height)];
            bytes[0] = (byte)'B';
            bytes[1] = (byte)'M';
            WriteInt32LittleEndian(bytes, 2, bytes.Length);
            WriteInt32LittleEndian(bytes, 10, pixelOffset);
            WriteInt32LittleEndian(bytes, 14, 40);
            WriteInt32LittleEndian(bytes, 18, width);
            WriteInt32LittleEndian(bytes, 22, height);
            WriteUInt16LittleEndian(bytes, 26, 1);
            WriteUInt16LittleEndian(bytes, 28, 32);

            for (int y = 0; y < height; y++) {
                int sourceY = height - 1 - y;
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

        private static byte[] CreateSinglePixelGif() =>
            Convert.FromBase64String("R0lGODlhAQABAJAAAAAAAP///ywAAAAAAQABAAACAkwBADs=");

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

        private static void AssertNoUnexpectedDiagnostics(
            IEnumerable<OfficeImageExportDiagnostic> diagnostics,
            params string[] expectedApproximationCodes) {
            // These tests qualify slide content and geometry, not host font cost or optional GSUB features.
            // Keep omissions, failures, and other approximations visible.
            OfficeImageExportDiagnostic[] unexpected = diagnostics.Where(diagnostic => {
                if ((diagnostic.Code == OfficeImageExportDiagnosticCodes.FontSubstituted || diagnostic.Code == "font-family-substitution")
                    && diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Warning
                    && diagnostic.LossKind == OfficeConversionLossKind.Approximation) return false;
                if (diagnostic.Code == "font-embedding-cost"
                    && diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Info
                    && diagnostic.LossKind == OfficeConversionLossKind.None) return false;
                if (diagnostic.Code == "unsupported-font-ligature-substitution"
                    && diagnostic.Source == "OfficeIMO.PowerPoint.Pdf"
                    && diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Warning
                    && diagnostic.LossKind == OfficeConversionLossKind.Approximation) return false;
                return !(expectedApproximationCodes.Contains(diagnostic.Code)
                    && diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Warning
                    && diagnostic.LossKind == OfficeConversionLossKind.Approximation);
            }).ToArray();
            Assert.Empty(unexpected);
        }
    }
}
