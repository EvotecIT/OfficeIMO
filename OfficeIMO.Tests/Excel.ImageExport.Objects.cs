using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using A = DocumentFormat.OpenXml.Drawing;
using X = DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using Threaded = DocumentFormat.OpenXml.Office2019.Excel.ThreadedComments;
using Xunit;

namespace OfficeIMO.Tests {
    public class ExcelImageExportObjectTests {
        [Fact]
        public void ExcelRange_ImageExportRendersCellCommentIndicatorsAndReportsUnsupportedBodiesInVisibleRange() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("Objects");
            sheet.CellValue(1, 1, "Name");
            sheet.CellValue(2, 2, "Reviewed");
            sheet.SetComment("B2", "Needs design review", "Reviewer");
            sheet.SetComment("C3", "Outside export", "Reviewer");

            ExcelRange range = sheet.Range("A1:B2");
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot();
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, new ExcelImageExportOptions { ShowGridlines = false });
            string svg = range.ToSvg(new ExcelImageExportOptions { ShowGridlines = false });

            OfficeImageExportDiagnostic snapshotDiagnostic = Assert.Single(snapshot.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.CellCommentUnsupported);
            OfficeImageExportDiagnostic pngDiagnostic = Assert.Single(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.CellCommentUnsupported);
            ExcelVisualCommentIndicator indicator = Assert.Single(snapshot.CommentIndicators);
            Assert.Equal(OfficeImageExportDiagnosticSeverity.Warning, snapshotDiagnostic.Severity);
            Assert.Equal("Objects!B2", snapshotDiagnostic.Source);
            Assert.Equal("Objects!B2", pngDiagnostic.Source);
            Assert.Equal("Objects!B2", indicator.Source);
            Assert.False(indicator.Threaded);
            Assert.Contains("<polygon", svg, StringComparison.Ordinal);
            Assert.Contains("#C00000", svg, StringComparison.Ordinal);
            Assert.DoesNotContain(snapshot.Diagnostics, diagnostic => diagnostic.Source == "Objects!C3");
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Source == "Objects!C3");
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            Assert.True(CountPixelsNear(rendered!, OfficeColor.FromRgb(192, 0, 0)) > 0);
        }

        [Fact]
        public void ExcelRange_ImageExportRendersCommentBodiesWhenRequested() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("CommentBody");
            sheet.CellValue(1, 1, "Name");
            sheet.CellValue(2, 2, "Reviewed");
            sheet.SetComment("B2", "Needs design review before this range is sent to leadership.", "Reviewer");

            ExcelImageExportOptions options = new ExcelImageExportOptions {
                ShowGridlines = false,
                ShowCommentBodies = true,
                DefaultColumnWidthPixels = 92D,
                DefaultRowHeightPixels = 28D
            };
            ExcelRange range = sheet.Range("A1:F8");
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            string svg = range.ToSvg(options);

            ExcelVisualCommentIndicator indicator = Assert.Single(snapshot.CommentIndicators);
            ExcelVisualCommentBody body = Assert.Single(snapshot.CommentBodies);
            ExcelVisualDrawingLayer commentLayer = Assert.Single(snapshot.DrawingLayers, layer => layer.Kind == ExcelVisualDrawingLayerKind.CommentBody);
            Assert.Equal(indicator.Source, body.Source);
            Assert.Same(body, commentLayer.CommentBody);
            Assert.Equal(0, commentLayer.Order);
            Assert.Equal(indicator.X + indicator.Width, body.AnchorX);
            Assert.True(body.AnchorX < body.X, "Comment body should be placed to the right of the annotated cell in this fixture.");
            Assert.Equal("Reviewer", body.Title);
            Assert.Contains("Needs design review", body.Text, StringComparison.Ordinal);
            Assert.Single(snapshot.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.CellCommentBodyApproximation);
            Assert.Single(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.CellCommentBodyApproximation);
            Assert.DoesNotContain(snapshot.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.CellCommentUnsupported);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.CellCommentUnsupported);
            Assert.Contains("Reviewer", svg, StringComparison.Ordinal);
            Assert.Contains("Needs design review", svg, StringComparison.Ordinal);
            Assert.Contains("font-weight=\"700\"", svg, StringComparison.Ordinal);
            Assert.Contains("text-anchor=\"start\"", svg, StringComparison.Ordinal);
            Assert.Contains("#FFFBE6", svg, StringComparison.Ordinal);
            Assert.Contains("#FFF2CC", svg, StringComparison.Ordinal);
            Assert.True(CountOccurrences(svg, "<polygon") >= 2, "Expected the comment indicator and anchored comment-body pointer to render as SVG polygons.");

            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            Assert.True(CountPixelsNear(rendered!, OfficeColor.FromRgb(255, 251, 230)) > 100);
            Assert.True(CountPixelsNear(rendered!, OfficeColor.FromRgb(255, 242, 204)) > 50);
            OfficeColor pointerPixel = rendered!.GetPixel((int)Math.Round((body.AnchorX + body.X) / 2D), (int)Math.Round(body.Y + 14D));
            Assert.True(
                pointerPixel.A >= 248 &&
                pointerPixel.R >= 190 &&
                pointerPixel.G >= 150 &&
                pointerPixel.B <= 240,
                $"Expected an anchored comment-body pointer pixel, but got {pointerPixel.A},{pointerPixel.R},{pointerPixel.G},{pointerPixel.B}.");
        }

        [Fact]
        public void ExcelRange_ImageExportRendersThreadedCommentIndicatorsAndReportsUnsupportedBodiesInVisibleRange() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            const string personId = "{11111111-1111-1111-1111-111111111111}";

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Threaded");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(1, 2, "Reviewed");
                sheet.CellValue(1, 4, "Outside");
                document.Save(false);
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                WorkbookPart workbookPart = spreadsheet.WorkbookPart!;
                WorkbookPersonPart personPart = workbookPart.AddNewPart<WorkbookPersonPart>();
                personPart.PersonList = new Threaded.PersonList(
                    new Threaded.Person {
                        DisplayName = "Modern Reviewer",
                        Id = personId,
                        UserId = "modern.reviewer@example.test",
                        ProviderId = "OfficeIMO.Tests"
                    });
                personPart.PersonList.Save();

                WorksheetPart worksheetPart = workbookPart.WorksheetParts.Single();
                WorksheetThreadedCommentsPart threadedPart = worksheetPart.AddNewPart<WorksheetThreadedCommentsPart>();
                threadedPart.ThreadedComments = new Threaded.ThreadedComments(
                    new Threaded.ThreadedComment(new Threaded.ThreadedCommentText("Review visible item")) {
                        Ref = "B1",
                        PersonId = personId,
                        Id = "{22222222-2222-2222-2222-222222222222}",
                        DT = new DateTime(2026, 6, 22, 10, 0, 0, DateTimeKind.Utc),
                    },
                    new Threaded.ThreadedComment(new Threaded.ThreadedCommentText("Outside exported range")) {
                        Ref = "D1",
                        PersonId = personId,
                        Id = "{33333333-3333-3333-3333-333333333333}",
                        DT = new DateTime(2026, 6, 22, 10, 5, 0, DateTimeKind.Utc),
                    });
                threadedPart.ThreadedComments.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                ExcelSheet sheet = document.Sheets.Single();
                ExcelRange range = sheet.Range("A1:B1");
                ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot();
                OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, new ExcelImageExportOptions { ShowGridlines = false });
                string svg = range.ToSvg(new ExcelImageExportOptions { ShowGridlines = false });

                OfficeImageExportDiagnostic snapshotDiagnostic = Assert.Single(snapshot.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.ThreadedCommentUnsupported);
                OfficeImageExportDiagnostic pngDiagnostic = Assert.Single(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.ThreadedCommentUnsupported);
                ExcelVisualCommentIndicator indicator = Assert.Single(snapshot.CommentIndicators);
                Assert.Equal(OfficeImageExportDiagnosticSeverity.Warning, snapshotDiagnostic.Severity);
                Assert.Equal("Threaded!B1", snapshotDiagnostic.Source);
                Assert.Equal("Threaded!B1", pngDiagnostic.Source);
                Assert.Equal("Threaded!B1", indicator.Source);
                Assert.True(indicator.Threaded);
                Assert.Contains("<polygon", svg, StringComparison.Ordinal);
                Assert.Contains("#7C3AED", svg, StringComparison.Ordinal);
                Assert.DoesNotContain(snapshot.Diagnostics, diagnostic => diagnostic.Source == "Threaded!D1");
                Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Source == "Threaded!D1");
                Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
                Assert.NotNull(rendered);
                Assert.True(CountPixelsNear(rendered!, OfficeColor.FromRgb(124, 58, 237)) > 0);
            }
        }

        [Fact]
        public void ExcelRange_ImageExportReportsUnsupportedDrawingShapesInVisibleRange() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Shapes");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(2, 2, "Reviewed");
                sheet.CellValue(1, 4, "Outside");
                document.Save(false);
            }

            AddDrawingShapes(filePath);

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                ExcelSheet sheet = document.Sheets.Single();
                ExcelRange range = sheet.Range("A1:B2");
                ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot();
                OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, new ExcelImageExportOptions { ShowGridlines = false });

                OfficeImageExportDiagnostic snapshotDiagnostic = Assert.Single(snapshot.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.DrawingShapeUnsupported);
                OfficeImageExportDiagnostic pngDiagnostic = Assert.Single(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.DrawingShapeUnsupported);
                Assert.Equal(OfficeImageExportDiagnosticSeverity.Warning, snapshotDiagnostic.Severity);
                Assert.Equal("Shapes!B2", snapshotDiagnostic.Source);
                Assert.Equal("Shapes!B2", pngDiagnostic.Source);
                Assert.Contains("Visible callout", snapshotDiagnostic.Message, StringComparison.Ordinal);
                Assert.DoesNotContain(snapshot.Diagnostics, diagnostic => diagnostic.Message.Contains("Outside callout", StringComparison.Ordinal));
                Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Message.Contains("Outside callout", StringComparison.Ordinal));
            }
        }

        [Fact]
        public void ExcelRange_ImageExportRendersSupportedDrawingShapesThroughSharedDrawing() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("ShapeVisual");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(2, 2, "Reviewed");
                document.Save(false);
            }

            AddSupportedDrawingShape(filePath);

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                ExcelSheet sheet = document.Sheets.Single();
                ExcelRange range = sheet.Range("A1:D4");
                var options = new ExcelImageExportOptions { ShowGridlines = false };
                ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
                OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
                string svg = range.ToSvg(options);

                Assert.True(
                    snapshot.DrawingObjects.Count == 1,
                    string.Join(" | ", snapshot.Diagnostics.Select(diagnostic => diagnostic.Code + ": " + diagnostic.Message)));
                ExcelVisualDrawingObject drawingObject = snapshot.DrawingObjects.Single();
                Assert.Equal("ShapeVisual!B2", drawingObject.Source);
                Assert.Equal("roundRect", drawingObject.ShapePresetName);
                Assert.Equal(OfficeShapeKind.RoundedRectangle, drawingObject.ShapeKind);
                Assert.Equal("Premium shape", drawingObject.Text);
                Assert.DoesNotContain(snapshot.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.DrawingShapeUnsupported);
                Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.DrawingShapeUnsupported);
                Assert.Contains("Premium shape", svg, StringComparison.Ordinal);
                Assert.Contains("#E0F2FE", svg, StringComparison.Ordinal);
                Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
                Assert.NotNull(rendered);
                Assert.True(CountPixelsNear(rendered!, OfficeColor.FromRgb(224, 242, 254)) > 100);
            }
        }

        [Fact]
        public void ExcelRange_ImageExportRendersSharedDrawingMlPresetShapesThroughSharedDrawing() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("PresetShape");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(2, 2, "Shared heart");
                document.Save(false);
            }

            AppendSupportedDrawingShape(
                filePath,
                "Shared heart",
                string.Empty,
                A.ShapeTypeValues.Heart,
                horizontalFlip: true,
                fillHex: "FECACA",
                strokeHex: "DC2626");

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                ExcelSheet sheet = document.Sheets.Single();
                ExcelRange range = sheet.Range("A1:D4");
                var options = new ExcelImageExportOptions { ShowGridlines = false };
                ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
                OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
                string svg = range.ToSvg(options);

                ExcelVisualDrawingObject drawingObject = Assert.Single(snapshot.DrawingObjects);
                Assert.Equal("PresetShape!B2", drawingObject.Source);
                Assert.Equal("heart", drawingObject.ShapePresetName);
                Assert.Equal(OfficeShapeKind.Path, drawingObject.ShapeKind);
                Assert.True(drawingObject.HorizontalFlip);
                Assert.False(drawingObject.VerticalFlip);
                Assert.DoesNotContain(snapshot.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.DrawingShapeUnsupported);
                Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.DrawingShapeUnsupported);
                Assert.Contains("<path", svg, StringComparison.Ordinal);
                Assert.Contains("#FECACA", svg, StringComparison.Ordinal);

                Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
                Assert.NotNull(rendered);
                Assert.True(CountPixelsNear(rendered!, OfficeColor.FromRgb(254, 202, 202)) > 100);
            }
        }

        [Fact]
        public void ExcelRange_ImageExportRendersRotatedDrawingMlPresetShapesThroughSharedDrawing() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("RotatedShape");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(2, 2, "Rotated heart");
                document.Save(false);
            }

            AppendSupportedDrawingShape(
                filePath,
                "Rotated heart",
                string.Empty,
                A.ShapeTypeValues.Heart,
                rotationDegrees: 35D,
                fillHex: "FDBA74",
                strokeHex: "EA580C");

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                ExcelSheet sheet = document.Sheets.Single();
                ExcelRange range = sheet.Range("A1:D4");
                var options = new ExcelImageExportOptions { ShowGridlines = false };
                ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
                OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
                OfficeImageExportResult svgResult = range.ExportImage(OfficeImageExportFormat.Svg, options);
                string svg = System.Text.Encoding.UTF8.GetString(svgResult.Bytes);

                ExcelVisualDrawingObject drawingObject = Assert.Single(snapshot.DrawingObjects);
                Assert.Equal("RotatedShape!B2", drawingObject.Source);
                Assert.Equal("heart", drawingObject.ShapePresetName);
                Assert.Equal(OfficeShapeKind.Path, drawingObject.ShapeKind);
                Assert.Equal(35D, drawingObject.RotationDegrees, precision: 3);
                Assert.True(drawingObject.HasRotation);
                Assert.DoesNotContain(snapshot.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.DrawingShapeUnsupported);
                Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.DrawingShapeUnsupported);
                Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.DrawingShapeTextRotationApproximation);
                Assert.DoesNotContain(svgResult.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.DrawingShapeTextRotationApproximation);
                Assert.Contains("<path", svg, StringComparison.Ordinal);
                Assert.Contains("matrix(", svg, StringComparison.Ordinal);
                Assert.Contains("#FDBA74", svg, StringComparison.Ordinal);

                Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
                Assert.NotNull(rendered);
                Assert.True(CountPixelsNear(rendered!, OfficeColor.FromRgb(253, 186, 116)) > 100);
            }
        }

        [Fact]
        public void ExcelRange_ImageExportReportsRotatedDrawingShapeTextApproximation() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("RotatedText");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(2, 2, "Rotated label");
                document.Save(false);
            }

            AppendSupportedDrawingShape(
                filePath,
                "Rotated label",
                "Rotated label",
                rotationDegrees: 25D);

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                ExcelSheet sheet = document.Sheets.Single();
                ExcelRange range = sheet.Range("A1:D4");
                var options = new ExcelImageExportOptions { ShowGridlines = false };
                ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
                OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
                OfficeImageExportResult svg = range.ExportImage(OfficeImageExportFormat.Svg, options);
                string svgText = System.Text.Encoding.UTF8.GetString(svg.Bytes);

                ExcelVisualDrawingObject drawingObject = Assert.Single(snapshot.DrawingObjects);
                Assert.Equal("RotatedText!B2", drawingObject.Source);
                Assert.Equal(25D, drawingObject.RotationDegrees, precision: 3);
                Assert.True(drawingObject.HasRotation);
                Assert.DoesNotContain(snapshot.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.DrawingShapeUnsupported);
                Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.DrawingShapeUnsupported);
                Assert.Single(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.DrawingShapeTextRotationApproximation);
                Assert.Single(svg.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.DrawingShapeTextRotationApproximation);
                Assert.Contains("Rotated label", svgText, StringComparison.Ordinal);
                Assert.Contains("transform=\"rotate(25", svgText, StringComparison.Ordinal);
            }
        }

        [Fact]
        public void ExcelRange_ImageExportHonorsDrawingShapeTextAlignmentThroughSharedDrawing() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("AlignedText");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(2, 2, "Aligned label");
                document.Save(false);
            }

            AppendSupportedDrawingShape(
                filePath,
                "Aligned label",
                "Aligned label",
                paragraphAlignment: A.TextAlignmentTypeValues.Right,
                verticalAlignment: A.TextAnchoringTypeValues.Bottom);

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                ExcelSheet sheet = document.Sheets.Single();
                ExcelRange range = sheet.Range("A1:D4");
                var options = new ExcelImageExportOptions { ShowGridlines = false };
                ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
                OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
                OfficeImageExportResult svg = range.ExportImage(OfficeImageExportFormat.Svg, options);
                string svgText = System.Text.Encoding.UTF8.GetString(svg.Bytes);

                ExcelVisualDrawingObject drawingObject = Assert.Single(snapshot.DrawingObjects);
                Assert.Equal("AlignedText!B2", drawingObject.Source);
                Assert.Equal(OfficeTextAlignment.Right, drawingObject.TextAlignment);
                Assert.Equal(OfficeTextVerticalAlignment.Bottom, drawingObject.TextVerticalAlignment);
                Assert.DoesNotContain(snapshot.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.DrawingShapeUnsupported);
                Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.DrawingShapeUnsupported);
                Assert.DoesNotContain(svg.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.DrawingShapeUnsupported);
                Assert.Contains("Aligned label", svgText, StringComparison.Ordinal);
                Assert.Contains("text-anchor=\"end\"", svgText, StringComparison.Ordinal);
            }
        }

        [Fact]
        public void ExcelRange_ImageExportHonorsDrawingShapeTextColorThroughSharedDrawing() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("TextColor");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(2, 2, "Colored label");
                document.Save(false);
            }

            AppendSupportedDrawingShape(
                filePath,
                "Colored label",
                "Colored label",
                fillHex: "FEF3C7",
                strokeHex: "D97706",
                textColorHex: "B91C1C");

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                ExcelSheet sheet = document.Sheets.Single();
                ExcelRange range = sheet.Range("A1:D4");
                var options = new ExcelImageExportOptions { ShowGridlines = false };
                ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
                OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
                OfficeImageExportResult svg = range.ExportImage(OfficeImageExportFormat.Svg, options);
                string svgText = System.Text.Encoding.UTF8.GetString(svg.Bytes);

                ExcelVisualDrawingObject drawingObject = Assert.Single(snapshot.DrawingObjects);
                Assert.Equal("FFB91C1C", drawingObject.TextColorArgb);
                Assert.Contains("fill=\"#B91C1C\"", svgText, StringComparison.Ordinal);
                Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
                Assert.NotNull(rendered);
                Assert.True(CountPixelsNear(rendered!, OfficeColor.FromRgb(185, 28, 28)) > 0);
            }
        }

        [Theory]
        [InlineData(true)]
        [InlineData(false)]
        public void ExcelRange_ImageExportPaintsShapesAndImagesInSourceDrawingLayerOrder(bool imageFirst) {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            byte[] redBadge = CreateSolidPng(120, 48, OfficeColor.FromRgb(220, 38, 38));
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("LayerOrder");
                sheet.CellValue(1, 1, "Layer order");
                sheet.CellValue(2, 2, "Overlap");
                if (imageFirst) {
                    sheet.AddImage(2, 2, redBadge, "image/png", widthPixels: 120, heightPixels: 48, name: "Red badge");
                }

                document.Save(false);
            }

            AppendSupportedDrawingShape(filePath, "Layer shape", string.Empty);
            if (!imageFirst) {
                using ExcelDocument document = ExcelDocument.Load(filePath);
                ExcelSheet sheet = document.Sheets.Single();
                sheet.AddImage(2, 2, redBadge, "image/png", widthPixels: 120, heightPixels: 48, name: "Red badge");
                document.Save(false);
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                ExcelSheet sheet = document.Sheets.Single();
                ExcelRange range = sheet.Range("A1:D4");
                var options = new ExcelImageExportOptions { ShowGridlines = false };
                ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
                string svg = range.ToSvg(options);
                OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);

                Assert.Equal(2, snapshot.DrawingLayers.Count);
                Assert.Equal(imageFirst ? ExcelVisualDrawingLayerKind.Image : ExcelVisualDrawingLayerKind.DrawingObject, snapshot.DrawingLayers[0].Kind);
                Assert.Equal(imageFirst ? ExcelVisualDrawingLayerKind.DrawingObject : ExcelVisualDrawingLayerKind.Image, snapshot.DrawingLayers[1].Kind);
                Assert.Equal(new[] { 0, 1 }, snapshot.DrawingLayers.Select(layer => layer.Order).ToArray());
                Assert.Single(snapshot.Images);
                Assert.Single(snapshot.DrawingObjects);

                int imageIndex = svg.IndexOf("<image", StringComparison.Ordinal);
                int shapeIndex = svg.IndexOf("#E0F2FE", StringComparison.Ordinal);
                Assert.True(imageIndex >= 0, svg);
                Assert.True(shapeIndex >= 0, svg);
                Assert.True(imageFirst ? imageIndex < shapeIndex : shapeIndex < imageIndex, "SVG drawing elements were not emitted in source layer order.");

                Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
                Assert.NotNull(rendered);
                ExcelVisualImage image = snapshot.Images.Single();
                ExcelVisualDrawingObject drawingObject = snapshot.DrawingObjects.Single();
                int sampleX = (int)Math.Round(Math.Max(image.X, drawingObject.X) + 20D);
                int sampleY = (int)Math.Round(Math.Max(image.Y, drawingObject.Y) + 20D);
                OfficeColor expected = imageFirst
                    ? OfficeColor.FromRgb(224, 242, 254)
                    : OfficeColor.FromRgb(220, 38, 38);
                AssertColorNear(rendered!.GetPixel(sampleX, sampleY), expected);
            }
        }

        private static void AddDrawingShapes(string filePath) {
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            DrawingsPart drawingsPart = worksheetPart.DrawingsPart ?? worksheetPart.AddNewPart<DrawingsPart>();
            if (worksheetPart.Worksheet!.Elements<X.Drawing>().FirstOrDefault() == null) {
                worksheetPart.Worksheet.Append(new X.Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });
            }

            drawingsPart.WorksheetDrawing = new Xdr.WorksheetDrawing(
                CreateShapeAnchor(1, 1, 2, 3, 2U, "Visible callout"),
                CreateShapeAnchor(3, 0, 4, 2, 3U, "Outside callout"));
            drawingsPart.WorksheetDrawing.Save();
            worksheetPart.Worksheet.Save();
        }

        private static void AddSupportedDrawingShape(string filePath) {
            AppendSupportedDrawingShape(filePath, "Premium shape", "Premium shape");
        }

        private static void AppendSupportedDrawingShape(
            string filePath,
            string name,
            string text,
            A.ShapeTypeValues? preset = null,
            bool horizontalFlip = false,
            bool verticalFlip = false,
            double rotationDegrees = 0D,
            string fillHex = "E0F2FE",
            string strokeHex = "0284C7",
            A.TextAlignmentTypeValues? paragraphAlignment = null,
            A.TextAnchoringTypeValues? verticalAlignment = null,
            string textColorHex = "1F2937") {
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            DrawingsPart drawingsPart = worksheetPart.DrawingsPart ?? worksheetPart.AddNewPart<DrawingsPart>();
            if (worksheetPart.Worksheet!.Elements<X.Drawing>().FirstOrDefault() == null) {
                worksheetPart.Worksheet.Append(new X.Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });
            }

            drawingsPart.WorksheetDrawing ??= new Xdr.WorksheetDrawing();
            drawingsPart.WorksheetDrawing.Append(
                CreateSupportedShapeAnchor(1, 1, 3, 3, 2U, name, text, preset, horizontalFlip, verticalFlip, rotationDegrees, fillHex, strokeHex, paragraphAlignment, verticalAlignment, textColorHex));
            drawingsPart.WorksheetDrawing.Save();
            worksheetPart.Worksheet.Save();
        }

        private static Xdr.TwoCellAnchor CreateShapeAnchor(int fromColumn, int fromRow, int toColumn, int toRow, uint id, string name) {
            return new Xdr.TwoCellAnchor(
                new Xdr.FromMarker(
                    new Xdr.ColumnId(fromColumn.ToString()),
                    new Xdr.ColumnOffset("0"),
                    new Xdr.RowId(fromRow.ToString()),
                    new Xdr.RowOffset("0")),
                new Xdr.ToMarker(
                    new Xdr.ColumnId(toColumn.ToString()),
                    new Xdr.ColumnOffset("0"),
                    new Xdr.RowId(toRow.ToString()),
                    new Xdr.RowOffset("0")),
                new Xdr.Shape(
                    new Xdr.NonVisualShapeProperties(
                        new Xdr.NonVisualDrawingProperties { Id = id, Name = name },
                        new Xdr.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true })),
                    new Xdr.ShapeProperties(new A.PresetGeometry { Preset = A.ShapeTypeValues.RoundRectangle }),
                    new Xdr.TextBody(
                        new A.BodyProperties(),
                        new A.ListStyle(),
                        new A.Paragraph(new A.Run(new A.Text(name))))),
                new Xdr.ClientData());
        }

        private static Xdr.TwoCellAnchor CreateSupportedShapeAnchor(
            int fromColumn,
            int fromRow,
            int toColumn,
            int toRow,
            uint id,
            string name,
            string? text = null,
            A.ShapeTypeValues? preset = null,
            bool horizontalFlip = false,
            bool verticalFlip = false,
            double rotationDegrees = 0D,
            string fillHex = "E0F2FE",
            string strokeHex = "0284C7",
            A.TextAlignmentTypeValues? paragraphAlignment = null,
            A.TextAnchoringTypeValues? verticalAlignment = null,
            string textColorHex = "1F2937") {
            var transform = new A.Transform2D {
                HorizontalFlip = horizontalFlip,
                VerticalFlip = verticalFlip
            };
            if (Math.Abs(rotationDegrees) > 0.0001D) {
                transform.Rotation = (int)Math.Round(rotationDegrees * 60000D);
            }

            var bodyProperties = new A.BodyProperties();
            if (verticalAlignment.HasValue) {
                bodyProperties.Anchor = verticalAlignment.Value;
            }

            var paragraph = new A.Paragraph();
            if (paragraphAlignment.HasValue) {
                paragraph.Append(new A.ParagraphProperties { Alignment = paragraphAlignment.Value });
            }

            paragraph.Append(new A.Run(
                new A.RunProperties(new A.SolidFill(new A.RgbColorModelHex { Val = textColorHex })),
                new A.Text(text ?? name)));

            return new Xdr.TwoCellAnchor(
                new Xdr.FromMarker(
                    new Xdr.ColumnId(fromColumn.ToString()),
                    new Xdr.ColumnOffset("0"),
                    new Xdr.RowId(fromRow.ToString()),
                    new Xdr.RowOffset("0")),
                new Xdr.ToMarker(
                    new Xdr.ColumnId(toColumn.ToString()),
                    new Xdr.ColumnOffset("0"),
                    new Xdr.RowId(toRow.ToString()),
                    new Xdr.RowOffset("0")),
                new Xdr.Shape(
                    new Xdr.NonVisualShapeProperties(
                        new Xdr.NonVisualDrawingProperties { Id = id, Name = name },
                        new Xdr.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true })),
                    new Xdr.ShapeProperties(
                        transform,
                        new A.PresetGeometry { Preset = preset ?? A.ShapeTypeValues.RoundRectangle },
                        new A.SolidFill(new A.RgbColorModelHex { Val = fillHex }),
                        new A.Outline(
                            new A.SolidFill(new A.RgbColorModelHex { Val = strokeHex })) {
                            Width = 12700
                        }),
                    new Xdr.TextBody(
                        bodyProperties,
                        new A.ListStyle(),
                        paragraph)),
                new Xdr.ClientData());
        }

        private static byte[] CreateSolidPng(int width, int height, OfficeColor color) {
            OfficeRasterImage image = new OfficeRasterImage(width, height, color);
            return OfficePngWriter.Encode(image);
        }

        private static void AssertColorNear(OfficeColor actual, OfficeColor expected) {
            Assert.True(
                Math.Abs(actual.R - expected.R) <= 8 &&
                Math.Abs(actual.G - expected.G) <= 8 &&
                Math.Abs(actual.B - expected.B) <= 8 &&
                actual.A >= 248,
                $"Expected ARGB near {expected.A},{expected.R},{expected.G},{expected.B} but got {actual.A},{actual.R},{actual.G},{actual.B}.");
        }

        private static int CountPixelsNear(OfficeRasterImage image, OfficeColor expected) {
            int count = 0;
            for (int y = 0; y < image.Height; y++) {
                for (int x = 0; x < image.Width; x++) {
                    OfficeColor color = image.GetPixel(x, y);
                    if (Math.Abs(color.R - expected.R) <= 8 &&
                        Math.Abs(color.G - expected.G) <= 8 &&
                        Math.Abs(color.B - expected.B) <= 8 &&
                        color.A >= 248) {
                        count++;
                    }
                }
            }

            return count;
        }

        private static int CountOccurrences(string text, string value) {
            int count = 0;
            int index = 0;
            while ((index = text.IndexOf(value, index, StringComparison.Ordinal)) >= 0) {
                count++;
                index += value.Length;
            }

            return count;
        }
    }
}
