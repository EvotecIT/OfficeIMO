using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.TestAssets;
using System.Threading;
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
            ExcelSheet sheet = document.AddWorksheet("Objects");
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
            ExcelSheet sheet = document.AddWorksheet("CommentBody");
            sheet.CellValue(1, 1, "Name");
            sheet.CellValue(2, 2, "Reviewed");
            sheet.SetCommentRichText("B2", new[] {
                new ExcelRichTextRun("Needs ") { FontColor = "C00000", Bold = true },
                new ExcelRichTextRun("design review") { Italic = true, Underline = true, FontColor = "2563EB" },
                new ExcelRichTextRun(" before this range is sent to leadership.")
            }, "Reviewer");

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
            Assert.Equal(3, body.RichTextRuns.Count);
            Assert.True(body.RichTextRuns[0].Bold);
            Assert.True(body.RichTextRuns[1].Italic);
            Assert.True(body.RichTextRuns[1].Underline);
            Assert.Single(snapshot.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.CellCommentBodyApproximation);
            Assert.Single(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.CellCommentBodyApproximation);
            Assert.DoesNotContain(snapshot.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.CellCommentUnsupported);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.CellCommentUnsupported);
            Assert.Contains("Reviewer", svg, StringComparison.Ordinal);
            Assert.Contains("Needs ", svg, StringComparison.Ordinal);
            Assert.Contains("design review", svg, StringComparison.Ordinal);
            Assert.Contains("font-weight=\"700\"", svg, StringComparison.Ordinal);
            Assert.Contains("font-style=\"italic\"", svg, StringComparison.Ordinal);
            Assert.Contains("text-decoration=\"underline\"", svg, StringComparison.Ordinal);
            Assert.Contains("#C00000", svg, StringComparison.Ordinal);
            Assert.Contains("#2563EB", svg, StringComparison.Ordinal);
            Assert.Contains("text-anchor=\"start\"", svg, StringComparison.Ordinal);
            Assert.Contains("#FFFBE6", svg, StringComparison.Ordinal);
            Assert.Contains("#FFF2CC", svg, StringComparison.Ordinal);
            Assert.True(CountOccurrences(svg, "<polygon") >= 2, "Expected the comment indicator and anchored comment-body pointer to render as SVG polygons.");

            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            Assert.True(CountPixelsNear(rendered!, OfficeColor.FromRgb(255, 251, 230)) > 100);
            Assert.True(CountPixelsNear(rendered!, OfficeColor.FromRgb(255, 242, 204)) > 50);
            Assert.True(CountBlueTextPixels(rendered!) > 0, "Expected the rich blue comment run to render into PNG output.");
            OfficeColor pointerPixel = rendered!.GetPixel((int)Math.Round((body.AnchorX + body.X) / 2D), (int)Math.Round(body.Y + 14D));
            Assert.True(
                pointerPixel.A >= 248 &&
                pointerPixel.R >= 190 &&
                pointerPixel.G >= 150 &&
                pointerPixel.B <= 240,
                $"Expected an anchored comment-body pointer pixel, but got {pointerPixel.A},{pointerPixel.R},{pointerPixel.G},{pointerPixel.B}.");
        }

        [Fact]
        public void ExcelRange_ImageExportPreservesRichCommentBreaksAndNonRgbColors() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("CommentBreaks");
                sheet.CellValue(2, 2, "Reviewed");
                sheet.SetCommentRichText("B2", new[] {
                    new ExcelRichTextRun("Line one"),
                    new ExcelRichTextRun("line two") { Bold = true }
                }, "Reviewer");
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
                X.Comment comment = worksheetPart.WorksheetCommentsPart!.Comments!.CommentList!.Elements<X.Comment>().Single();
                List<X.Run> runs = comment.CommentText!.Elements<X.Run>().ToList();
                comment.CommentText.InsertAfter(new X.Break(), runs[0]);
                X.Color color = runs[1].RunProperties!.GetFirstChild<X.Color>() ?? runs[1].RunProperties!.AppendChild(new X.Color());
                color.Rgb = null;
                color.Indexed = 5U;
                color.Tint = -0.25D;
                worksheetPart.WorksheetCommentsPart.Comments!.Save();
            }

            using ExcelDocument loadedDocument = ExcelDocument.Load(filePath);
            ExcelSheet loadedSheet = loadedDocument.Sheets.Single();
            ExcelRange range = loadedSheet.Range("A1:F8");
            var options = new ExcelImageExportOptions {
                ShowGridlines = false,
                ShowCommentBodies = true,
                DefaultColumnWidthPixels = 92D,
                DefaultRowHeightPixels = 28D
            };

            ExcelVisualCommentBody body = Assert.Single(range.CreateVisualSnapshot(options).CommentBodies);
            string svg = range.ToSvg(options);

            Assert.Equal("Line one\nline two", body.Text);
            Assert.Equal("Line one\n", string.Concat(body.RichTextRuns.Take(2).Select(run => run.Text)));
            Assert.Equal("line two", body.RichTextRuns[2].Text);
            Assert.True(body.RichTextRuns[2].Bold);
            Assert.StartsWith("FF", body.RichTextRuns[2].FontColorArgb, StringComparison.Ordinal);
            string themedRunColor = "#" + body.RichTextRuns[2].FontColorArgb!.Substring(2);
            Assert.NotEqual("#1F2937", themedRunColor);
            Assert.Contains("Line one", svg, StringComparison.Ordinal);
            Assert.Contains("line two", svg, StringComparison.Ordinal);
            Assert.Contains(themedRunColor, svg, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void ExcelRange_ImageExportPlacesCommentBodyAwayFromChartWhenSpaceExists() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("CommentChart");
            for (int column = 1; column <= 10; column++) {
                sheet.SetColumnWidth(column, 10);
            }

            for (int row = 1; row <= 8; row++) {
                sheet.SetRowHeight(row, 28);
            }

            sheet.CellValue(4, 1, "Month");
            sheet.CellValue(4, 2, "Score");
            sheet.CellValue(5, 1, "Jan");
            sheet.CellValue(5, 2, 120);
            sheet.CellValue(6, 1, "Feb");
            sheet.CellValue(6, 2, 180);
            sheet.CellValue(2, 4, "Review");
            sheet.SetComment("D2", "This note should avoid the chart when rendered.", "Reviewer");
            sheet.AddChartFromRange("A4:B6", row: 1, column: 5, widthPixels: 260, heightPixels: 120, type: ExcelChartType.ColumnClustered, title: "Trend");

            var options = new ExcelImageExportOptions {
                ShowGridlines = false,
                ShowCommentBodies = true,
                DefaultColumnWidthPixels = 70D,
                DefaultRowHeightPixels = 28D
            };
            ExcelRangeVisualSnapshot snapshot = sheet.Range("A1:J8").CreateVisualSnapshot(options);
            OfficeImageExportResult svg = sheet.Range("A1:J8").ExportImage(OfficeImageExportFormat.Svg, options);

            ExcelVisualCommentIndicator indicator = Assert.Single(snapshot.CommentIndicators);
            ExcelVisualCommentBody body = Assert.Single(snapshot.CommentBodies);
            ExcelVisualChart chart = Assert.Single(snapshot.Charts);
            Assert.True(body.X < indicator.X, "The right-side placement overlaps the chart, so the body should move to available space on the left.");
            Assert.False(Intersects(body.X, body.Y, body.Width, body.Height, chart.X, chart.Y, chart.Width, chart.Height));
            Assert.Single(svg.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.CellCommentBodyApproximation);
        }

        [Fact]
        public void ExcelRange_ImageExportRendersThreadedCommentIndicatorsAndReportsUnsupportedBodiesInVisibleRange() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            const string personId = "{11111111-1111-1111-1111-111111111111}";

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Threaded");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(1, 2, "Reviewed");
                sheet.CellValue(1, 4, "Outside");
                document.Save();
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
                ExcelSheet sheet = document.AddWorksheet("Shapes");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(2, 2, "Reviewed");
                sheet.CellValue(1, 4, "Outside");
                document.Save();
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
                ExcelSheet sheet = document.AddWorksheet("ShapeVisual");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(2, 2, "Reviewed");
                document.Save();
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
        public void ExcelRange_ImageExportCarriesShapingContextIntoDrawingObjects() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("ShapeShaping");
                sheet.CellValue(1, 1, "Name");
                document.Save();
            }

            AppendSupportedDrawingShape(
                filePath,
                "Shaped object",
                "A",
                textFontFamily: ManagedTextShapingTestAssets.FamilyName);

            using ExcelDocument loaded = ExcelDocument.Load(filePath);
            ExcelRange range = loaded.Sheets.Single().Range("A1:D4");
            var provider = new ManagedTextShapingTestAssets.RecordingProvider();
            var options = new ExcelImageExportOptions {
                ShowGridlines = false,
                TextShapingProvider = provider,
                TextShapingLanguage = "ar-SA"
            };
            options.Fonts.Add(
                ManagedTextShapingTestAssets.FamilyName,
                ManagedTextShapingTestAssets.CreateFont('A'));
            using var cancellation = new CancellationTokenSource();

            OfficeImageExportResult result =
                range.ExportImage(OfficeImageExportFormat.Png, options, cancellation.Token);

            Assert.Contains(provider.Requests, request =>
                request.Text == "A" &&
                request.Language == "ar-SA" &&
                request.CancellationToken == cancellation.Token);
            Assert.DoesNotContain(
                result.Diagnostics,
                diagnostic => diagnostic.Code == OfficeImageExportDiagnosticCodes.TextShapingFallback);
        }

        [Fact]
        public void ExcelWorksheet_ImageExportExpandsUsedRangeForDrawingShapes() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("ShapeOnly");
                sheet.CellValue(1, 1, "Only cell");
                document.Save();
            }

            AppendSupportedDrawingShape(filePath, "Outside used range", "Outside used range");

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                ExcelSheet sheet = document.Sheets.Single();
                var options = new ExcelWorksheetImageExportOptions { ShowGridlines = false };
                ExcelRangeVisualSnapshot snapshot = sheet.CreateVisualSnapshot(options);
                OfficeImageExportResult png = sheet.ExportImage(OfficeImageExportFormat.Png, options);
                OfficeImageExportResult svg = sheet.ExportImage(OfficeImageExportFormat.Svg, options);
                string svgText = System.Text.Encoding.UTF8.GetString(svg.Bytes);

                Assert.True(
                    snapshot.DrawingObjects.Count == 1,
                    string.Join(" | ", snapshot.Diagnostics.Select(diagnostic => diagnostic.Code + ": " + diagnostic.Message)));
                ExcelVisualDrawingObject drawingObject = snapshot.DrawingObjects.Single();
                Assert.Equal("ShapeOnly!B2", drawingObject.Source);
                Assert.Equal("Outside used range", drawingObject.Text);
                Assert.Equal("ShapeOnly!A1:C3", png.Source);
                Assert.Equal(png.Source, svg.Source);
                Assert.Contains("#E0F2FE", svgText, StringComparison.Ordinal);
                Assert.DoesNotContain(snapshot.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.DrawingShapeUnsupported);
                Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.DrawingShapeUnsupported);
            }
        }

        [Fact]
        public void ExcelWorksheet_ImageExportDoesNotExpandUsedRangeForUnsupportedDrawingShapes() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("UnsupportedOnly");
                sheet.CellValue(1, 1, "Only cell");
                document.Save();
            }

            AddDrawingShapes(filePath);

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                ExcelSheet sheet = document.Sheets.Single();
                var options = new ExcelWorksheetImageExportOptions { ShowGridlines = false };
                ExcelRangeVisualSnapshot snapshot = sheet.CreateVisualSnapshot(options);
                OfficeImageExportResult png = sheet.ExportImage(OfficeImageExportFormat.Png, options);

                Assert.Equal("UnsupportedOnly!A1:A1", png.Source);
                Assert.Empty(snapshot.DrawingObjects);
                Assert.DoesNotContain(snapshot.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.DrawingShapeUnsupported);
                Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.DrawingShapeUnsupported);
            }
        }

        [Fact]
        public void ExcelWorksheet_ImageExportIncludesDrawingObjectAnchorOffsetsWhenExpandingUsedRange() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("OffsetShape");
                sheet.SetColumnWidth(1, 8D);
                sheet.SetColumnWidth(2, 8D);
                sheet.SetColumnWidth(3, 8D);
                sheet.CellValue(1, 1, "Only cell");
                document.Save();
            }

            AppendSupportedDrawingShape(filePath, "Offset shape", "Offset shape", offsetXPixels: 90, toColumn: 3, toRow: 3);

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                ExcelSheet sheet = document.Sheets.Single();
                var options = new ExcelWorksheetImageExportOptions { ShowGridlines = false };
                OfficeImageExportResult png = sheet.ExportImage(OfficeImageExportFormat.Png, options);

                Assert.Equal("OffsetShape!A1:C3", png.Source);
                Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.DrawingShapeUnsupported);
            }
        }

        [Fact]
        public void ExcelRange_ImageExportResolvesAbsoluteAnchorDrawingShapeCoordinates() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("AbsoluteShape");
                sheet.SetColumnWidth(1, 8);
                sheet.SetColumnWidth(2, 8);
                sheet.SetRowHeight(1, 24);
                sheet.CellValue(1, 2, "Shape area");
                document.Save();
            }

            AppendAbsoluteDrawingShape(filePath, "Absolute shape", "Absolute shape", xPixels: 80, yPixels: 12, widthPixels: 96, heightPixels: 34);

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                ExcelSheet sheet = document.Sheets.Single();
                ExcelRangeVisualSnapshot snapshot = sheet.Range("B1:D3").CreateVisualSnapshot(new ExcelImageExportOptions { ShowGridlines = false });

                ExcelVisualDrawingObject drawingObject = Assert.Single(snapshot.DrawingObjects);
                Assert.Equal("Absolute shape", drawingObject.Text);
                Assert.True(drawingObject.X > 0D, "Absolute-anchor shapes should resolve from worksheet-canvas coordinates instead of falling back to A1.");
                Assert.Equal(19D, drawingObject.X);
                Assert.True(drawingObject.Y > 0D, "Absolute-anchor shapes should preserve their vertical worksheet-canvas offset.");
                Assert.Equal(96D, drawingObject.Width);
                Assert.Equal(34D, drawingObject.Height);
            }
        }

        [Fact]
        public void ExcelRange_ImageExportResolvesThemedDrawingShapeColorsThroughSharedDrawing() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                document.EnsureWorkbookTheme();
                ExcelSheet sheet = document.AddWorksheet("ThemeShape");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(2, 2, "Themed shape");
                document.Save();
            }

            AppendSupportedDrawingShape(
                filePath,
                "Themed shape",
                "Theme text",
                fillSchemeColor: "accent1",
                fillLuminanceModulation: 60000,
                fillLuminanceOffset: 40000,
                strokeSchemeColor: "accent2",
                textSchemeColor: "accent3");

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                ExcelSheet sheet = document.Sheets.Single();
                ExcelRange range = sheet.Range("A1:D4");
                var options = new ExcelImageExportOptions { ShowGridlines = false };
                ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
                OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
                string svg = range.ToSvg(options);

                ExcelVisualDrawingObject drawingObject = Assert.Single(snapshot.DrawingObjects);
                Assert.Equal("FF96B3D7", drawingObject.FillColorArgb);
                Assert.Equal("FFC0504D", drawingObject.StrokeColorArgb);
                Assert.Equal("FF9BBB59", drawingObject.TextColorArgb);
                Assert.DoesNotContain(snapshot.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.DrawingShapeUnsupported);
                Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.DrawingShapeUnsupported);
                Assert.Contains("#96B3D7", svg, StringComparison.Ordinal);
                Assert.Contains("#C0504D", svg, StringComparison.Ordinal);
                Assert.Contains("#9BBB59", svg, StringComparison.Ordinal);

                Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
                Assert.NotNull(rendered);
                Assert.True(CountPixelsNear(rendered!, OfficeColor.FromRgb(149, 179, 215)) > 100);
            }
        }

        [Fact]
        public void ExcelRange_ImageExportHonorsDrawingShapeOutlineStyleThroughSharedDrawing() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("OutlineStyle");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(2, 2, "Styled outline");
                document.Save();
            }

            AppendSupportedDrawingShape(
                filePath,
                "Styled outline",
                string.Empty,
                preset: A.ShapeTypeValues.Rectangle,
                strokeDash: A.PresetLineDashValues.DashDot,
                strokeLineCap: A.LineCapValues.Round,
                strokeLineJoin: OfficeStrokeLineJoin.Bevel);

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                ExcelSheet sheet = document.Sheets.Single();
                ExcelRange range = sheet.Range("A1:D4");
                var options = new ExcelImageExportOptions { ShowGridlines = false };
                ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
                OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
                string svg = range.ToSvg(options);

                ExcelVisualDrawingObject drawingObject = Assert.Single(snapshot.DrawingObjects);
                Assert.Equal(OfficeStrokeDashStyle.DashDot, drawingObject.StrokeDashStyle);
                Assert.Equal(OfficeStrokeLineCap.Round, drawingObject.StrokeLineCap);
                Assert.Equal(OfficeStrokeLineJoin.Bevel, drawingObject.StrokeLineJoin);
                Assert.DoesNotContain(snapshot.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.DrawingShapeUnsupported);
                Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.DrawingShapeUnsupported);
                Assert.Contains("stroke-dasharray", svg, StringComparison.Ordinal);
                Assert.Contains("stroke-linecap=\"round\"", svg, StringComparison.Ordinal);
                Assert.Contains("stroke-linejoin=\"bevel\"", svg, StringComparison.Ordinal);
            }
        }

        [Fact]
        public void ExcelRange_ImageExportRendersDrawingShapeEffectsThroughSharedDrawing() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("ShapeEffects");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(2, 2, "Styled effects");
                document.Save();
            }

            AppendSupportedDrawingShape(
                filePath,
                "Effect shape",
                string.Empty,
                preset: A.ShapeTypeValues.Rectangle,
                fillHex: "FDE68A",
                strokeHex: "B45309",
                glowHex: "2563EB",
                glowAlpha: 55000,
                glowRadiusPixels: 8,
                shadowHex: "111827",
                shadowAlpha: 45000,
                shadowDistancePixels: 7,
                shadowBlurPixels: 5,
                shadowDirectionDegrees: 45);

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                ExcelSheet sheet = document.Sheets.Single();
                ExcelRange range = sheet.Range("A1:D4");
                var options = new ExcelImageExportOptions { ShowGridlines = false };
                ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
                OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
                string svg = range.ToSvg(options);

                ExcelVisualDrawingObject drawingObject = Assert.Single(snapshot.DrawingObjects);
                Assert.NotNull(drawingObject.Glow);
                Assert.NotNull(drawingObject.Shadow);
                Assert.Equal(8D, drawingObject.Glow!.Radius);
                Assert.Equal(5D, drawingObject.Shadow!.BlurRadius);
                Assert.True(drawingObject.Shadow.OffsetX > 0D);
                Assert.True(drawingObject.Shadow.OffsetY > 0D);
                Assert.DoesNotContain(snapshot.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.DrawingShapeUnsupported);
                Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.DrawingShapeUnsupported);
                Assert.Contains("#2563EB", svg, StringComparison.Ordinal);
                Assert.Contains("#111827", svg, StringComparison.Ordinal);

                Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
                Assert.NotNull(rendered);
                Assert.True(CountPixelsNear(rendered!, OfficeColor.FromRgb(37, 99, 235)) > 0, "Expected the blue glow to paint into PNG output.");
                Assert.True(CountDarkPixels(rendered!) > 0, "Expected the dark shadow to paint into PNG output.");
            }
        }

        [Fact]
        public void ExcelRange_ImageExportPreservesDrawingShapeTextLineBreaks() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("ShapeBreaks");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(2, 2, "Shape breaks");
                document.Save();
            }

            AppendSupportedDrawingShape(filePath, "Break shape", "Line 1", textHardBreak: true);

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                ExcelSheet sheet = document.Sheets.Single();
                ExcelRangeVisualSnapshot snapshot = sheet.Range("A1:D4").CreateVisualSnapshot(new ExcelImageExportOptions { ShowGridlines = false });

                ExcelVisualDrawingObject drawingObject = Assert.Single(snapshot.DrawingObjects);
                Assert.Equal("Line 1" + Environment.NewLine + "Line 2", drawingObject.Text);
            }
        }

        [Fact]
        public void ExcelRange_ImageExportRendersDefaultStyledDrawingShapeFill() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                document.EnsureWorkbookTheme();
                ExcelSheet sheet = document.AddWorksheet("DefaultStyle");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(2, 2, "Default styled shape");
                document.Save();
            }

            AppendDefaultStyledDrawingShape(filePath, "Default styled shape", "Styled shape");

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                ExcelSheet sheet = document.Sheets.Single();
                ExcelRange range = sheet.Range("A1:D4");
                var options = new ExcelImageExportOptions { ShowGridlines = false };
                ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
                OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);

                ExcelVisualDrawingObject drawingObject = Assert.Single(snapshot.DrawingObjects);
                Assert.False(string.IsNullOrWhiteSpace(drawingObject.FillColorArgb));
                Assert.False(string.IsNullOrWhiteSpace(drawingObject.StrokeColorArgb));
                Assert.True(drawingObject.StrokeWidth > 0D);
                Assert.DoesNotContain(snapshot.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.DrawingShapeUnsupported);
                Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.DrawingShapeUnsupported);
            }
        }

        [Fact]
        public void ExcelRange_ImageExportRendersSharedDrawingMlPresetShapesThroughSharedDrawing() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("PresetShape");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(2, 2, "Shared heart");
                document.Save();
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
        public void ExcelRange_ImageExportRendersAdditionalFlowchartPresetShapesThroughSharedDrawing() {
            (A.ShapeTypeValues Preset, string ExpectedName)[] presets = {
                (A.ShapeTypeValues.FlowChartPredefinedProcess, "flowChartPredefinedProcess"),
                (A.ShapeTypeValues.FlowChartInternalStorage, "flowChartInternalStorage"),
                (A.ShapeTypeValues.FlowChartMagneticDisk, "flowChartMagneticDisk"),
                (A.ShapeTypeValues.FlowChartMagneticTape, "flowChartMagneticTape"),
                (A.ShapeTypeValues.FlowChartMagneticDrum, "flowChartMagneticDrum"),
                (A.ShapeTypeValues.FlowChartMultidocument, "flowChartMultidocument"),
                (A.ShapeTypeValues.FlowChartPunchedTape, "flowChartPunchedTape"),
                (A.ShapeTypeValues.FlowChartSummingJunction, "flowChartSummingJunction"),
                (A.ShapeTypeValues.FlowChartSort, "flowChartSort"),
                (A.ShapeTypeValues.FlowChartDisplay, "flowChartDisplay")
            };

            foreach ((A.ShapeTypeValues preset, string expectedName) in presets) {
                string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
                using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                    ExcelSheet sheet = document.AddWorksheet("FlowchartShape");
                    sheet.CellValue(1, 1, "Name");
                    sheet.CellValue(2, 2, expectedName);
                    document.Save();
                }

                AppendSupportedDrawingShape(
                    filePath,
                    expectedName,
                    string.Empty,
                    preset,
                    fillHex: "E0F2FE",
                    strokeHex: "0284C7");

                using ExcelDocument loaded = ExcelDocument.Load(filePath);
                ExcelSheet loadedSheet = loaded.Sheets.Single();
                ExcelRange range = loadedSheet.Range("A1:D4");
                var options = new ExcelImageExportOptions { ShowGridlines = false };
                ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
                OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
                string svg = range.ToSvg(options);

                ExcelVisualDrawingObject drawingObject = Assert.Single(snapshot.DrawingObjects);
                Assert.Equal(expectedName, drawingObject.ShapePresetName);
                Assert.Equal(OfficeShapeKind.Path, drawingObject.ShapeKind);
                Assert.DoesNotContain(snapshot.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.DrawingShapeUnsupported);
                Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.DrawingShapeUnsupported);
                Assert.Contains("<path", svg, StringComparison.Ordinal);
            }
        }

        [Fact]
        public void ExcelRange_ImageExportRendersRotatedDrawingMlPresetShapesThroughSharedDrawing() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("RotatedShape");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(2, 2, "Rotated heart");
                document.Save();
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
                ExcelSheet sheet = document.AddWorksheet("RotatedText");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(2, 2, "Rotated label");
                document.Save();
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
                ExcelSheet sheet = document.AddWorksheet("AlignedText");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(2, 2, "Aligned label");
                document.Save();
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
                ExcelSheet sheet = document.AddWorksheet("TextColor");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(2, 2, "Colored label");
                document.Save();
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

        [Fact]
        public void ExcelRange_ImageExportHonorsDrawingShapeTextFontThroughSharedDrawing() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("TextFont");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(2, 2, "Styled label");
                document.Save();
            }

            AppendSupportedDrawingShape(
                filePath,
                "Styled label",
                "Styled label",
                fillHex: "DBEAFE",
                strokeHex: "2563EB",
                textColorHex: "111827",
                textFontFamily: "Aptos",
                textFontSize: 18D,
                textBold: true,
                textItalic: true,
                textUnderline: true);

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                ExcelSheet sheet = document.Sheets.Single();
                ExcelRange range = sheet.Range("A1:D4");
                var options = new ExcelImageExportOptions { ShowGridlines = false };
                ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
                OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
                OfficeImageExportResult svg = range.ExportImage(OfficeImageExportFormat.Svg, options);
                string svgText = System.Text.Encoding.UTF8.GetString(svg.Bytes);

                ExcelVisualDrawingObject drawingObject = Assert.Single(snapshot.DrawingObjects);
                Assert.Equal("Aptos", drawingObject.TextFontFamily);
                Assert.Equal(18D, drawingObject.TextFontSize);
                Assert.True((drawingObject.TextFontStyle & OfficeFontStyle.Bold) == OfficeFontStyle.Bold);
                Assert.True((drawingObject.TextFontStyle & OfficeFontStyle.Italic) == OfficeFontStyle.Italic);
                Assert.True((drawingObject.TextFontStyle & OfficeFontStyle.Underline) == OfficeFontStyle.Underline);
                Assert.Contains("font-family=\"Aptos\"", svgText, StringComparison.Ordinal);
                Assert.Contains("font-size=\"18\"", svgText, StringComparison.Ordinal);
                Assert.Contains("font-weight=\"700\"", svgText, StringComparison.Ordinal);
                Assert.Contains("font-style=\"italic\"", svgText, StringComparison.Ordinal);
                Assert.Contains("text-decoration=\"underline\"", svgText, StringComparison.Ordinal);
                Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
                Assert.NotNull(rendered);
                Assert.True(CountPixelsNear(rendered!, OfficeColor.FromRgb(17, 24, 39)) > 0);
            }
        }

        [Fact]
        public void ExcelRange_ImageExportWrapsDrawingShapeTextThroughSharedDrawing() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                document.AddWorksheet("TextWrap");
                document.Save();
            }

            AppendSupportedDrawingShape(
                filePath,
                "Wrapped label",
                "Alpha beta gamma delta epsilon",
                fillHex: "EFF6FF",
                strokeHex: "2563EB",
                paragraphAlignment: A.TextAlignmentTypeValues.Left,
                verticalAlignment: A.TextAnchoringTypeValues.Top,
                textColorHex: "111827",
                textFontSize: 8D,
                textWrap: true);

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                ExcelSheet sheet = document.Sheets.Single();
                ExcelRange range = sheet.Range("A1:D4");
                var options = new ExcelImageExportOptions { ShowGridlines = false };
                ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
                OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
                OfficeImageExportResult svg = range.ExportImage(OfficeImageExportFormat.Svg, options);
                string svgText = System.Text.Encoding.UTF8.GetString(svg.Bytes);

                ExcelVisualDrawingObject drawingObject = Assert.Single(snapshot.DrawingObjects);
                Assert.True(drawingObject.TextWrap);
                Assert.Equal(10D, drawingObject.TextInsetLeft);
                Assert.Equal(5D, drawingObject.TextInsetTop);
                Assert.Equal(10D, drawingObject.TextInsetRight);
                Assert.Equal(5D, drawingObject.TextInsetBottom);
                Assert.Contains("Alpha", svgText, StringComparison.Ordinal);
                Assert.Contains("epsilon", svgText, StringComparison.Ordinal);
                Assert.True(CountOccurrences(svgText, "<text") >= 2);
                Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
                Assert.NotNull(rendered);
                Assert.True(CountDarkPixels(rendered!) > 0);
            }
        }

        [Fact]
        public void ExcelRange_ImageExportHonorsDrawingShapeTextInsetsThroughSharedDrawing() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                document.AddWorksheet("TextInsets");
                document.Save();
            }

            AppendSupportedDrawingShape(
                filePath,
                "Inset label",
                "Inset label",
                fillHex: "F8FAFC",
                strokeHex: "475569",
                paragraphAlignment: A.TextAlignmentTypeValues.Left,
                verticalAlignment: A.TextAnchoringTypeValues.Top,
                textColorHex: "111827",
                textFontSize: 12D,
                textInsetLeftEmu: 0,
                textInsetTopEmu: 0,
                textInsetRightEmu: 0,
                textInsetBottomEmu: 0);

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                ExcelSheet sheet = document.Sheets.Single();
                ExcelRange range = sheet.Range("A1:D4");
                var options = new ExcelImageExportOptions { ShowGridlines = false };
                ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
                OfficeImageExportResult svg = range.ExportImage(OfficeImageExportFormat.Svg, options);
                string svgText = System.Text.Encoding.UTF8.GetString(svg.Bytes);

                ExcelVisualDrawingObject drawingObject = Assert.Single(snapshot.DrawingObjects);
                Assert.Equal(0D, drawingObject.TextInsetLeft);
                Assert.Equal(0D, drawingObject.TextInsetTop);
                Assert.Equal(0D, drawingObject.TextInsetRight);
                Assert.Equal(0D, drawingObject.TextInsetBottom);
                Assert.Contains("Inset label", svgText, StringComparison.Ordinal);
                Assert.Contains("x=\"0\"", svgText, StringComparison.Ordinal);
            }
        }

        [Fact]
        public void ExcelRange_ImageExportShrinksDrawingShapeTextThroughSharedDrawing() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                document.AddWorksheet("TextShrink");
                document.Save();
            }

            AppendSupportedDrawingShape(
                filePath,
                "Shrink label",
                "Alpha beta gamma delta epsilon zeta",
                fillHex: "F1F5F9",
                strokeHex: "334155",
                paragraphAlignment: A.TextAlignmentTypeValues.Left,
                verticalAlignment: A.TextAnchoringTypeValues.Top,
                textColorHex: "111827",
                textFontSize: 24D,
                textWrap: true,
                textShrinkToFit: true);

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                ExcelSheet sheet = document.Sheets.Single();
                ExcelRange range = sheet.Range("A1:D4");
                var options = new ExcelImageExportOptions { ShowGridlines = false };
                ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
                OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
                OfficeImageExportResult svg = range.ExportImage(OfficeImageExportFormat.Svg, options);
                string svgText = System.Text.Encoding.UTF8.GetString(svg.Bytes);

                ExcelVisualDrawingObject drawingObject = Assert.Single(snapshot.DrawingObjects);
                double fontSize = ExtractFirstSvgFontSize(svgText);
                Assert.True(drawingObject.TextShrinkToFit);
                Assert.True(fontSize < 24D, "Expected DrawingML normalAutoFit to shrink the rendered SVG font size.");
                Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
                Assert.NotNull(rendered);
                Assert.True(CountDarkPixels(rendered!) > 0);
            }
        }

        [Fact]
        public void ExcelRange_ImageExportReportsDrawingShapeTextAutoFitUnsupported() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                document.AddWorksheet("TextAutoFit");
                document.Save();
            }

            AppendSupportedDrawingShape(
                filePath,
                "AutoFit label",
                "Resize this shape to fit me",
                fillHex: "F8FAFC",
                strokeHex: "475569",
                paragraphAlignment: A.TextAlignmentTypeValues.Left,
                verticalAlignment: A.TextAnchoringTypeValues.Top,
                textColorHex: "111827",
                textFontSize: 14D,
                textWrap: true,
                textResizeShapeToFit: true);

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                ExcelSheet sheet = document.Sheets.Single();
                ExcelRange range = sheet.Range("A1:D4");
                var options = new ExcelImageExportOptions { ShowGridlines = false };
                ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
                OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
                OfficeImageExportResult svg = range.ExportImage(OfficeImageExportFormat.Svg, options);
                string svgText = System.Text.Encoding.UTF8.GetString(svg.Bytes);

                ExcelVisualDrawingObject drawingObject = Assert.Single(snapshot.DrawingObjects);
                Assert.True(drawingObject.TextResizeShapeToFit);
                Assert.DoesNotContain(snapshot.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.DrawingShapeTextAutoFitUnsupported);
                Assert.Single(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.DrawingShapeTextAutoFitUnsupported && diagnostic.Source == "TextAutoFit!B2");
                Assert.Single(svg.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.DrawingShapeTextAutoFitUnsupported && diagnostic.Source == "TextAutoFit!B2");
                Assert.Contains("Resize", svgText, StringComparison.Ordinal);
            }
        }

        [Fact]
        public void ExcelRange_ImageExportRendersSimpleDrawingShapeVerticalTextThroughSharedDrawing() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                document.AddWorksheet("VerticalText");
                document.Save();
            }

            AppendSupportedDrawingShape(
                filePath,
                "Vertical label",
                "Vertical label",
                fillHex: "F8FAFC",
                strokeHex: "475569",
                paragraphAlignment: A.TextAlignmentTypeValues.Left,
                verticalAlignment: A.TextAnchoringTypeValues.Top,
                textColorHex: "111827",
                textFontSize: 14D,
                textOrientation: A.TextVerticalValues.Vertical);

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                ExcelSheet sheet = document.Sheets.Single();
                ExcelRange range = sheet.Range("A1:D4");
                var options = new ExcelImageExportOptions { ShowGridlines = false };
                ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
                OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
                OfficeImageExportResult svg = range.ExportImage(OfficeImageExportFormat.Svg, options);
                string svgText = System.Text.Encoding.UTF8.GetString(svg.Bytes);

                ExcelVisualDrawingObject drawingObject = Assert.Single(snapshot.DrawingObjects);
                Assert.Equal(ExcelDrawingTextOrientation.Vertical, drawingObject.TextOrientation);
                Assert.DoesNotContain(snapshot.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.DrawingShapeTextVerticalOrientationUnsupported);
                Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.DrawingShapeTextVerticalOrientationUnsupported);
                Assert.DoesNotContain(svg.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.DrawingShapeTextVerticalOrientationUnsupported);
                Assert.Equal("Vertical label", drawingObject.Text);
                Assert.Contains("<text", svgText, StringComparison.Ordinal);
                Assert.DoesNotContain(">Vertical label</text>", svgText, StringComparison.Ordinal);
                Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
                Assert.NotNull(rendered);
                Assert.True(CountDarkPixels(rendered!) > 0);
            }
        }

        [Fact]
        public void ExcelRange_ImageExportRendersVertical270DrawingShapeTextThroughSharedDrawing() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                document.AddWorksheet("Vertical270");
                document.Save();
            }

            AppendSupportedDrawingShape(
                filePath,
                "Vertical 270 label",
                "V270",
                fillHex: "F8FAFC",
                strokeHex: "475569",
                paragraphAlignment: A.TextAlignmentTypeValues.Left,
                verticalAlignment: A.TextAnchoringTypeValues.Top,
                textColorHex: "111827",
                textFontSize: 14D,
                textOrientation: A.TextVerticalValues.Vertical270,
                toColumn: 4,
                toRow: 6);

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                ExcelSheet sheet = document.Sheets.Single();
                ExcelRange range = sheet.Range("A1:F8");
                var options = new ExcelImageExportOptions { ShowGridlines = false };
                ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
                OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
                OfficeImageExportResult svg = range.ExportImage(OfficeImageExportFormat.Svg, options);
                string svgText = System.Text.Encoding.UTF8.GetString(svg.Bytes);

                ExcelVisualDrawingObject drawingObject = Assert.Single(snapshot.DrawingObjects);
                Assert.Equal(ExcelDrawingTextOrientation.Vertical270, drawingObject.TextOrientation);
                Assert.DoesNotContain(snapshot.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.DrawingShapeTextVerticalOrientationUnsupported);
                Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.DrawingShapeTextVerticalOrientationUnsupported);
                Assert.DoesNotContain(svg.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.DrawingShapeTextVerticalOrientationUnsupported);
                Assert.Single(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.DrawingShapeTextRotationApproximation && diagnostic.Source == "Vertical270!B2");
                Assert.Single(svg.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.DrawingShapeTextRotationApproximation && diagnostic.Source == "Vertical270!B2");
                Assert.Contains("V270", svgText, StringComparison.Ordinal);
                Assert.Contains("transform=\"rotate(270", svgText, StringComparison.Ordinal);
                Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
                Assert.NotNull(rendered);
            }
        }

        [Theory]
        [InlineData(true)]
        [InlineData(false)]
        public void ExcelRange_ImageExportPaintsShapesAndImagesInSourceDrawingLayerOrder(bool imageFirst) {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            byte[] redBadge = CreateSolidPng(120, 48, OfficeColor.FromRgb(220, 38, 38));
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("LayerOrder");
                sheet.CellValue(1, 1, "Layer order");
                sheet.CellValue(2, 2, "Overlap");
                if (imageFirst) {
                    sheet.AddImage(2, 2, redBadge, "image/png", widthPixels: 120, heightPixels: 48, name: "Red badge");
                }

                document.Save();
            }

            AppendSupportedDrawingShape(filePath, "Layer shape", string.Empty);
            if (!imageFirst) {
                using ExcelDocument document = ExcelDocument.Load(filePath);
                ExcelSheet sheet = document.Sheets.Single();
                sheet.AddImage(2, 2, redBadge, "image/png", widthPixels: 120, heightPixels: 48, name: "Red badge");
                document.Save();
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
            string textColorHex = "1F2937",
            string? fillSchemeColor = null,
            int? fillLuminanceModulation = null,
            int? fillLuminanceOffset = null,
            string? strokeSchemeColor = null,
            string? textSchemeColor = null,
            string? textFontFamily = null,
            double? textFontSize = null,
            bool textBold = false,
            bool textItalic = false,
            bool textUnderline = false,
            bool textWrap = false,
            bool textShrinkToFit = false,
            bool textResizeShapeToFit = false,
            A.TextVerticalValues? textOrientation = null,
            int? textInsetLeftEmu = null,
            int? textInsetTopEmu = null,
            int? textInsetRightEmu = null,
            int? textInsetBottomEmu = null,
            bool textHardBreak = false,
            int toColumn = 3,
            int toRow = 3,
            int offsetXPixels = 0,
            int offsetYPixels = 0,
            A.PresetLineDashValues? strokeDash = null,
            A.LineCapValues? strokeLineCap = null,
            OfficeStrokeLineJoin? strokeLineJoin = null,
            string? glowHex = null,
            int? glowAlpha = null,
            int glowRadiusPixels = 0,
            string? shadowHex = null,
            int? shadowAlpha = null,
            int shadowDistancePixels = 0,
            int shadowBlurPixels = 0,
            int shadowDirectionDegrees = 0) {
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            DrawingsPart drawingsPart = worksheetPart.DrawingsPart ?? worksheetPart.AddNewPart<DrawingsPart>();
            if (worksheetPart.Worksheet!.Elements<X.Drawing>().FirstOrDefault() == null) {
                worksheetPart.Worksheet.Append(new X.Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });
            }

            drawingsPart.WorksheetDrawing ??= new Xdr.WorksheetDrawing();
            drawingsPart.WorksheetDrawing.Append(
                CreateSupportedShapeAnchor(1, 1, toColumn, toRow, 2U, name, text, preset, horizontalFlip, verticalFlip, rotationDegrees, fillHex, strokeHex, paragraphAlignment, verticalAlignment, textColorHex, fillSchemeColor, fillLuminanceModulation, fillLuminanceOffset, strokeSchemeColor, textSchemeColor, textFontFamily, textFontSize, textBold, textItalic, textUnderline, textWrap, textShrinkToFit, textResizeShapeToFit, textOrientation, textInsetLeftEmu, textInsetTopEmu, textInsetRightEmu, textInsetBottomEmu, textHardBreak, offsetXPixels, offsetYPixels, strokeDash, strokeLineCap, strokeLineJoin, glowHex, glowAlpha, glowRadiusPixels, shadowHex, shadowAlpha, shadowDistancePixels, shadowBlurPixels, shadowDirectionDegrees));
            drawingsPart.WorksheetDrawing.Save();
            worksheetPart.Worksheet.Save();
        }

        private static void AppendDefaultStyledDrawingShape(string filePath, string name, string text) {
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            DrawingsPart drawingsPart = worksheetPart.DrawingsPart ?? worksheetPart.AddNewPart<DrawingsPart>();
            if (worksheetPart.Worksheet!.Elements<X.Drawing>().FirstOrDefault() == null) {
                worksheetPart.Worksheet.Append(new X.Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });
            }

            drawingsPart.WorksheetDrawing ??= new Xdr.WorksheetDrawing();
            drawingsPart.WorksheetDrawing.Append(new Xdr.TwoCellAnchor(
                new Xdr.FromMarker(
                    new Xdr.ColumnId("1"),
                    new Xdr.ColumnOffset("0"),
                    new Xdr.RowId("1"),
                    new Xdr.RowOffset("0")),
                new Xdr.ToMarker(
                    new Xdr.ColumnId("3"),
                    new Xdr.ColumnOffset("0"),
                    new Xdr.RowId("3"),
                    new Xdr.RowOffset("0")),
                new Xdr.Shape(
                    new Xdr.NonVisualShapeProperties(
                        new Xdr.NonVisualDrawingProperties { Id = 77U, Name = name },
                        new Xdr.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true })),
                    new Xdr.ShapeProperties(
                        new A.PresetGeometry { Preset = A.ShapeTypeValues.Rectangle }),
                    new Xdr.ShapeStyle(
                        new A.LineReference(new A.SchemeColor { Val = A.SchemeColorValues.Accent2 }) { Index = 2U },
                        new A.FillReference(new A.SchemeColor { Val = A.SchemeColorValues.Accent4 }) { Index = 1U },
                        new A.EffectReference(new A.SchemeColor { Val = A.SchemeColorValues.Accent4 }) { Index = 0U },
                        new A.FontReference(new A.SchemeColor { Val = A.SchemeColorValues.Text1 }) { Index = A.FontCollectionIndexValues.Minor }),
                    new Xdr.TextBody(
                        new A.BodyProperties(),
                        new A.ListStyle(),
                        new A.Paragraph(new A.Run(new A.Text(text))))),
                new Xdr.ClientData()));
            drawingsPart.WorksheetDrawing.Save();
            worksheetPart.Worksheet.Save();
        }

        private static void AppendAbsoluteDrawingShape(string filePath, string name, string text, int xPixels, int yPixels, int widthPixels, int heightPixels) {
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            DrawingsPart drawingsPart = worksheetPart.DrawingsPart ?? worksheetPart.AddNewPart<DrawingsPart>();
            if (worksheetPart.Worksheet!.Elements<X.Drawing>().FirstOrDefault() == null) {
                worksheetPart.Worksheet.Append(new X.Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });
            }

            drawingsPart.WorksheetDrawing ??= new Xdr.WorksheetDrawing();
            drawingsPart.WorksheetDrawing.Append(new Xdr.AbsoluteAnchor(
                new Xdr.Position { X = xPixels * 9525L, Y = yPixels * 9525L },
                new Xdr.Extent { Cx = widthPixels * 9525L, Cy = heightPixels * 9525L },
                new Xdr.Shape(
                    new Xdr.NonVisualShapeProperties(
                        new Xdr.NonVisualDrawingProperties { Id = 88U, Name = name },
                        new Xdr.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true })),
                    new Xdr.ShapeProperties(
                        new A.Transform2D(
                            new A.Offset { X = 0L, Y = 0L },
                            new A.Extents { Cx = widthPixels * 9525L, Cy = heightPixels * 9525L }),
                        new A.PresetGeometry { Preset = A.ShapeTypeValues.RoundRectangle },
                        CreateSolidFill("E0F2FE"),
                        new A.Outline(CreateSolidFill("0284C7")) { Width = 12700 }),
                    new Xdr.TextBody(
                        new A.BodyProperties(),
                        new A.ListStyle(),
                        new A.Paragraph(new A.Run(
                            new A.RunProperties(),
                            new A.Text(text))))),
                new Xdr.ClientData()));
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
            string textColorHex = "1F2937",
            string? fillSchemeColor = null,
            int? fillLuminanceModulation = null,
            int? fillLuminanceOffset = null,
            string? strokeSchemeColor = null,
            string? textSchemeColor = null,
            string? textFontFamily = null,
            double? textFontSize = null,
            bool textBold = false,
            bool textItalic = false,
            bool textUnderline = false,
            bool textWrap = false,
            bool textShrinkToFit = false,
            bool textResizeShapeToFit = false,
            A.TextVerticalValues? textOrientation = null,
            int? textInsetLeftEmu = null,
            int? textInsetTopEmu = null,
            int? textInsetRightEmu = null,
            int? textInsetBottomEmu = null,
            bool textHardBreak = false,
            int offsetXPixels = 0,
            int offsetYPixels = 0,
            A.PresetLineDashValues? strokeDash = null,
            A.LineCapValues? strokeLineCap = null,
            OfficeStrokeLineJoin? strokeLineJoin = null,
            string? glowHex = null,
            int? glowAlpha = null,
            int glowRadiusPixels = 0,
            string? shadowHex = null,
            int? shadowAlpha = null,
            int shadowDistancePixels = 0,
            int shadowBlurPixels = 0,
            int shadowDirectionDegrees = 0) {
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

            if (textOrientation.HasValue) {
                bodyProperties.Vertical = textOrientation.Value;
            }

            if (textWrap) {
                bodyProperties.Wrap = A.TextWrappingValues.Square;
            }

            if (textShrinkToFit) {
                bodyProperties.Append(new A.NormalAutoFit());
            }

            if (textResizeShapeToFit) {
                bodyProperties.Append(new A.ShapeAutoFit());
            }

            if (textInsetLeftEmu.HasValue) {
                bodyProperties.LeftInset = textInsetLeftEmu.Value;
            }

            if (textInsetTopEmu.HasValue) {
                bodyProperties.TopInset = textInsetTopEmu.Value;
            }

            if (textInsetRightEmu.HasValue) {
                bodyProperties.RightInset = textInsetRightEmu.Value;
            }

            if (textInsetBottomEmu.HasValue) {
                bodyProperties.BottomInset = textInsetBottomEmu.Value;
            }

            var paragraph = new A.Paragraph();
            if (paragraphAlignment.HasValue) {
                paragraph.Append(new A.ParagraphProperties { Alignment = paragraphAlignment.Value });
            }

            var runProperties = new A.RunProperties {
                Bold = textBold,
                Italic = textItalic
            };
            if (textFontSize.HasValue) {
                runProperties.FontSize = (int)Math.Round(textFontSize.Value * 100D);
            }

            if (textUnderline) {
                runProperties.Underline = A.TextUnderlineValues.Single;
            }

            runProperties.Append(CreateSolidFill(textColorHex, textSchemeColor));
            if (!string.IsNullOrWhiteSpace(textFontFamily)) {
                runProperties.Append(new A.LatinFont { Typeface = textFontFamily });
            }

            if (textHardBreak) {
                paragraph.Append(
                    new A.Run(
                        (A.RunProperties)runProperties.CloneNode(true),
                        new A.Text(text ?? name)),
                    new A.Break(),
                    new A.Run(
                        (A.RunProperties)runProperties.CloneNode(true),
                        new A.Text("Line 2")));
            } else {
                paragraph.Append(new A.Run(
                    runProperties,
                    new A.Text(text ?? name)));
            }

            var outline = new A.Outline(
                CreateSolidFill(strokeHex, strokeSchemeColor)) {
                Width = 12700
            };
            if (strokeDash.HasValue) {
                outline.Append(new A.PresetDash { Val = strokeDash.Value });
            }

            if (strokeLineCap.HasValue) {
                outline.CapType = strokeLineCap.Value;
            }

            if (strokeLineJoin.HasValue) {
                switch (strokeLineJoin.Value) {
                    case OfficeStrokeLineJoin.Round:
                        outline.Append(new A.Round());
                        break;
                    case OfficeStrokeLineJoin.Bevel:
                        outline.Append(new A.Bevel());
                        break;
                    case OfficeStrokeLineJoin.Miter:
                        outline.Append(new A.Miter());
                        break;
                }
            }

            A.EffectList? effects = CreateShapeEffects(glowHex, glowAlpha, glowRadiusPixels, shadowHex, shadowAlpha, shadowDistancePixels, shadowBlurPixels, shadowDirectionDegrees);
            var shapeProperties = new Xdr.ShapeProperties(
                transform,
                new A.PresetGeometry { Preset = preset ?? A.ShapeTypeValues.RoundRectangle },
                CreateSolidFill(fillHex, fillSchemeColor, fillLuminanceModulation, fillLuminanceOffset),
                outline);
            if (effects != null) {
                shapeProperties.Append(effects);
            }

            return new Xdr.TwoCellAnchor(
                new Xdr.FromMarker(
                    new Xdr.ColumnId(fromColumn.ToString()),
                    new Xdr.ColumnOffset(((long)offsetXPixels * 9525L).ToString(System.Globalization.CultureInfo.InvariantCulture)),
                    new Xdr.RowId(fromRow.ToString()),
                    new Xdr.RowOffset(((long)offsetYPixels * 9525L).ToString(System.Globalization.CultureInfo.InvariantCulture))),
                new Xdr.ToMarker(
                    new Xdr.ColumnId(toColumn.ToString()),
                    new Xdr.ColumnOffset("0"),
                    new Xdr.RowId(toRow.ToString()),
                    new Xdr.RowOffset("0")),
                new Xdr.Shape(
                    new Xdr.NonVisualShapeProperties(
                        new Xdr.NonVisualDrawingProperties { Id = id, Name = name },
                        new Xdr.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true })),
                    shapeProperties,
                    new Xdr.TextBody(
                        bodyProperties,
                        new A.ListStyle(),
                        paragraph)),
                new Xdr.ClientData());
        }

        private static A.EffectList? CreateShapeEffects(
            string? glowHex,
            int? glowAlpha,
            int glowRadiusPixels,
            string? shadowHex,
            int? shadowAlpha,
            int shadowDistancePixels,
            int shadowBlurPixels,
            int shadowDirectionDegrees) {
            var effects = new A.EffectList();
            if (!string.IsNullOrWhiteSpace(glowHex) && glowRadiusPixels > 0) {
                var glow = new A.Glow { Radius = glowRadiusPixels * 9525L };
                glow.Append(CreateEffectColor(glowHex!, glowAlpha));
                effects.Append(glow);
            }

            if (!string.IsNullOrWhiteSpace(shadowHex) && (shadowDistancePixels > 0 || shadowBlurPixels > 0)) {
                var shadow = new A.OuterShadow {
                    BlurRadius = shadowBlurPixels * 9525L,
                    Distance = shadowDistancePixels * 9525L,
                    Direction = shadowDirectionDegrees * 60000
                };
                shadow.Append(CreateEffectColor(shadowHex!, shadowAlpha));
                effects.Append(shadow);
            }

            return effects.HasChildren ? effects : null;
        }

        private static A.RgbColorModelHex CreateEffectColor(string rgbHex, int? alpha) {
            var color = new A.RgbColorModelHex { Val = rgbHex };
            if (alpha.HasValue) {
                color.Append(new A.Alpha { Val = alpha.Value });
            }

            return color;
        }

        private static A.SolidFill CreateSolidFill(string rgbHex, string? schemeColor = null, int? luminanceModulation = null, int? luminanceOffset = null) {
            var fill = new A.SolidFill();
            OpenXmlCompositeElement color = string.IsNullOrWhiteSpace(schemeColor)
                ? new A.RgbColorModelHex { Val = rgbHex }
                : new A.SchemeColor { Val = ResolveSchemeColor(schemeColor!) };
            if (luminanceModulation.HasValue) {
                color.Append(new A.LuminanceModulation { Val = luminanceModulation.Value });
            }

            if (luminanceOffset.HasValue) {
                color.Append(new A.LuminanceOffset { Val = luminanceOffset.Value });
            }

            fill.Append(color);
            return fill;
        }

        private static A.SchemeColorValues ResolveSchemeColor(string value) =>
            value switch {
                "accent1" => A.SchemeColorValues.Accent1,
                "accent2" => A.SchemeColorValues.Accent2,
                "accent3" => A.SchemeColorValues.Accent3,
                "accent4" => A.SchemeColorValues.Accent4,
                "accent5" => A.SchemeColorValues.Accent5,
                "accent6" => A.SchemeColorValues.Accent6,
                "tx1" => A.SchemeColorValues.Text1,
                "tx2" => A.SchemeColorValues.Text2,
                "bg1" => A.SchemeColorValues.Background1,
                "bg2" => A.SchemeColorValues.Background2,
                "hlink" => A.SchemeColorValues.Hyperlink,
                "folHlink" => A.SchemeColorValues.FollowedHyperlink,
                _ => throw new ArgumentOutOfRangeException(nameof(value), value, "Unsupported DrawingML scheme color fixture value.")
            };

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

        private static int CountBlueTextPixels(OfficeRasterImage image) {
            int count = 0;
            for (int y = 0; y < image.Height; y++) {
                for (int x = 0; x < image.Width; x++) {
                    OfficeColor color = image.GetPixel(x, y);
                    if (color.A >= 248 &&
                        color.B > color.R + 35 &&
                        color.B > color.G + 15 &&
                        color.R < 120 &&
                        color.G < 170) {
                        count++;
                    }
                }
            }

            return count;
        }

        private static int CountDarkPixels(OfficeRasterImage image) {
            int count = 0;
            for (int y = 0; y < image.Height; y++) {
                for (int x = 0; x < image.Width; x++) {
                    OfficeColor color = image.GetPixel(x, y);
                    if (color.A >= 248 && color.R < 90 && color.G < 100 && color.B < 120) {
                        count++;
                    }
                }
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

        private static int CountOccurrences(string text, string value) {
            int count = 0;
            int index = 0;
            while ((index = text.IndexOf(value, index, StringComparison.Ordinal)) >= 0) {
                count++;
                index += value.Length;
            }

            return count;
        }

        private static bool Intersects(
            double firstX,
            double firstY,
            double firstWidth,
            double firstHeight,
            double secondX,
            double secondY,
            double secondWidth,
            double secondHeight) =>
            firstX < secondX + secondWidth &&
            firstX + firstWidth > secondX &&
            firstY < secondY + secondHeight &&
            firstY + firstHeight > secondY;
    }
}
