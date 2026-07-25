using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Html;
using OfficeIMO.PowerPoint.Pdf;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Tests {
    public class PowerPointSharedBubbleCharts {
        [Fact]
        public void BubbleChart_CreateUpdateReopen_PreservesNativeDataAndWorkbook() {
            string output = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".pptx");
            OfficeChartData initial = CreateBubbleData(
                new[] { 2D, 4D, 6D }, new[] { 10D, 20D, 30D }, new[] { 9D, 25D, 49D },
                new OfficeColor?[] {
                    OfficeColor.Parse("#D62828"),
                    OfficeColor.Parse("#2A9D8F"),
                    OfficeColor.Parse("#F4A261")
                });
            OfficeChartData updated = CreateBubbleData(
                new[] { 3D, 6D, 9D }, new[] { 12D, 24D, 36D }, new[] { 16D, 36D, 64D });
            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(output)) {
                    PowerPointChart chart = presentation.AddSlide().AddChartPoints(
                        OfficeChartKind.Bubble, initial, 36, 30, 420, 250,
                        new PowerPointChartAccessibilityOptions {
                            AlternativeText = "Portfolio position by return, risk, and investment"
                        });

                    Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot authored));
                    Assert.Equal(OfficeChartKind.Bubble, authored.ChartKind);
                    Assert.Equal(new[] { 9D, 25D, 49D },
                        Assert.Single(authored.Data.Series).BubbleSizes);
                    Assert.Equal(OfficeColor.Parse("#2A9D8F"),
                        Assert.Single(authored.Data.Series).PointColors![1]);
                    Assert.Contains("Series\tX\tY\tSize", chart.CreateDataSummary(),
                        StringComparison.Ordinal);
                    Assert.Contains("Portfolio\t6\t30\t49", chart.CreateDataSummary(),
                        StringComparison.Ordinal);

                    chart.SetSeriesFillColor(0, "123456")
                        .SetSeriesLineColor(0, "654321", widthPoints: 2.5D)
                        .SetDataLabels(showValue: true)
                        .SetSeriesDataLabels(0, showValue: true, numberFormat: "0.0");
                    ChartPart authoredChartPart = presentation.Slides[0].SlidePart.ChartParts.Single();
                    C.BubbleChart authoredBubble = authoredChartPart.ChartSpace!
                        .GetFirstChild<C.Chart>()!.GetFirstChild<C.PlotArea>()!
                        .GetFirstChild<C.BubbleChart>()!;
                    authoredBubble.GetFirstChild<C.BubbleScale>()!.Val = 145U;
                    authoredBubble.GetFirstChild<C.ShowNegativeBubbles>()!.Val = true;
                    authoredBubble.GetFirstChild<C.SizeRepresents>()!.Val =
                        C.SizeRepresentsValues.Width;
                    authoredChartPart.ChartSpace.Save();

                    chart.UpdateData(updated);
                    Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot refreshed));
                    Assert.Equal(new[] { 3D, 6D, 9D },
                        Assert.Single(refreshed.Data.Series).XValues);
                    Assert.Equal(new[] { 12D, 24D, 36D },
                        Assert.Single(refreshed.Data.Series).Values);
                    Assert.Equal(new[] { 16D, 36D, 64D },
                        Assert.Single(refreshed.Data.Series).BubbleSizes);
                    Assert.Equal(OfficeColor.Parse("#D62828"),
                        Assert.Single(refreshed.Data.Series).PointColors![0]);
                    Assert.Equal(2.5D,
                        Assert.Single(refreshed.Data.Series).MarkerOutlineWidth);
                    Assert.Equal(OfficeColor.Parse("#654321"),
                        Assert.Single(refreshed.Data.Series).MarkerOutlineColor);
                    Assert.Equal(145D, refreshed.BubbleScalePercent);
                    Assert.Equal(OfficeChartBubbleSizeMode.Width,
                        refreshed.BubbleSizeMode);
                    Assert.True(chart.TryGetSnapshot(
                        out PowerPointChartSnapshot nativeSnapshot));
                    OfficeChartSnapshot pdfSnapshot =
                        PowerPointPdfConverterExtensions.CreateOfficeChartSnapshot(
                            nativeSnapshot, 420D, 250D,
                            new PowerPointPdfSaveOptions());
                    Assert.Equal(145D, pdfSnapshot.BubbleScalePercent);
                    Assert.Equal(OfficeChartBubbleSizeMode.Width,
                        pdfSnapshot.BubbleSizeMode);
                    Assert.Equal(OfficeColor.Parse("#654321"),
                        Assert.Single(pdfSnapshot.Data.Series).MarkerOutlineColor);
                    Assert.Equal(2.5D,
                        Assert.Single(pdfSnapshot.Data.Series).MarkerOutlineWidth);
                    C.BubbleChart preservedBubble = authoredChartPart.ChartSpace!
                        .GetFirstChild<C.Chart>()!.GetFirstChild<C.PlotArea>()!
                        .GetFirstChild<C.BubbleChart>()!;
                    Assert.Equal(145U,
                        preservedBubble.GetFirstChild<C.BubbleScale>()!.Val!.Value);
                    Assert.True(preservedBubble.GetFirstChild<C.ShowNegativeBubbles>()!
                        .Val!.Value);
                    Assert.Equal(C.SizeRepresentsValues.Width,
                        preservedBubble.GetFirstChild<C.SizeRepresents>()!.Val!.Value);
                    Assert.NotNull(preservedBubble.GetFirstChild<C.DataLabels>());
                    C.BubbleChartSeries preservedSeries =
                        Assert.Single(preservedBubble.Elements<C.BubbleChartSeries>());
                    Assert.NotNull(preservedSeries.GetFirstChild<C.DataLabels>());
                    Assert.Equal("123456", preservedSeries.GetFirstChild<C.ChartShapeProperties>()!
                        .GetFirstChild<DocumentFormat.OpenXml.Drawing.SolidFill>()!
                        .RgbColorModelHex!.Val!.Value);

                    List<ValidationErrorInfo> validation = presentation.ValidateDocument();
                    Assert.True(validation.Count == 0, string.Join(Environment.NewLine,
                        validation.Select(error =>
                            (error.Path?.XPath ?? string.Empty) + ": " + error.Description)));
                    presentation.Save();
                }

                using (PowerPointPresentation reopened = PowerPointPresentation.Load(output,
                           new PowerPointLoadOptions {
                               AccessMode = DocumentAccessMode.ReadOnly
                           })) {
                    PowerPointChart chart = Assert.Single(reopened.Slides[0].Charts);
                    Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
                    Assert.Equal(OfficeChartKind.Bubble, snapshot.ChartKind);
                    OfficeChartSeries series = Assert.Single(snapshot.Data.Series);
                    Assert.Equal(new[] { 3D, 6D, 9D }, series.XValues);
                    Assert.Equal(new[] { 12D, 24D, 36D }, series.Values);
                    Assert.Equal(new[] { 16D, 36D, 64D }, series.BubbleSizes);
                    Assert.Equal(OfficeColor.Parse("#F4A261"), series.PointColors![2]);
                    Assert.Equal(2.5D, series.MarkerOutlineWidth);
                    Assert.Equal(OfficeColor.Parse("#654321"), series.MarkerOutlineColor);
                    Assert.Equal(145D, snapshot.BubbleScalePercent);
                    Assert.Equal(OfficeChartBubbleSizeMode.Width,
                        snapshot.BubbleSizeMode);
                }

                using PresentationDocument document = PresentationDocument.Open(output, false);
                ChartPart chartPart = Assert.Single(document.PresentationPart!.SlideParts
                    .SelectMany(slidePart => slidePart.ChartParts));
                C.PlotArea plotArea = chartPart.ChartSpace!.GetFirstChild<C.Chart>()!
                    .GetFirstChild<C.PlotArea>()!;
                C.BubbleChart bubbleChart = Assert.Single(plotArea.Elements<C.BubbleChart>());
                C.BubbleChartSeries nativeSeries =
                    Assert.Single(bubbleChart.Elements<C.BubbleChartSeries>());
                C.BubbleSize bubbleSize = nativeSeries.GetFirstChild<C.BubbleSize>()!;
                Assert.Contains("$C$2:$C$4",
                    bubbleSize.NumberReference!.Formula!.Text,
                    StringComparison.Ordinal);
                Assert.Equal("64", bubbleSize.NumberReference.NumberingCache!
                    .Elements<C.NumericPoint>().Last().NumericValue!.Text);
                Assert.Equal(2, plotArea.Elements<C.ValueAxis>().Count());

                EmbeddedPackagePart package =
                    Assert.Single(chartPart.GetPartsOfType<EmbeddedPackagePart>());
                using SpreadsheetDocument workbook =
                    SpreadsheetDocument.Open(package.GetStream(), false);
                S.Cell sizeCell = workbook.WorkbookPart!.WorksheetParts.Single().Worksheet
                    .Descendants<S.Cell>().Single(cell => cell.CellReference?.Value == "C4");
                Assert.Equal("64", sizeCell.CellValue?.Text);
                Assert.Empty(new OpenXmlValidator().Validate(chartPart));
            } finally {
                if (File.Exists(output)) File.Delete(output);
            }
        }

        [Fact]
        public void BubbleChart_AuthorsSeriesOutlineFromBubbleOptions() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create(new MemoryStream());
            OfficeColor outlineColor = OfficeColor.Parse("#AABBCC");
            var data = new OfficeChartData(new[] { "1" }, new[] {
                OfficeChartSeries.CreateBubble("Portfolio",
                    new[] { 1D },
                    new[] { 2D },
                    new[] { 4D },
                    color: OfficeColor.Parse("#112233"),
                    markerOutlineColor: outlineColor,
                    markerOutlineWidth: 1.5D)
            });

            PowerPointChart chart = presentation.AddSlide().AddChart(
                OfficeChartKind.Bubble, data);
            C.BubbleChartSeries nativeSeries = presentation.Slides[0].SlidePart
                .ChartParts.Single().ChartSpace!
                .Descendants<C.BubbleChartSeries>().Single();
            DocumentFormat.OpenXml.Drawing.Outline outline = nativeSeries
                .GetFirstChild<C.ChartShapeProperties>()!
                .GetFirstChild<DocumentFormat.OpenXml.Drawing.Outline>()!;

            Assert.Equal(19050, outline.Width!.Value);
            Assert.Equal("AABBCC", outline
                .GetFirstChild<DocumentFormat.OpenXml.Drawing.SolidFill>()!
                .RgbColorModelHex!.Val!.Value);
            Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
            Assert.Equal(outlineColor,
                Assert.Single(snapshot.Data.Series).MarkerOutlineColor);
        }

        [Fact]
        public void BubbleChart_PreservesAlphaForSeriesPointAndOutlineColors() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create(new MemoryStream());
            OfficeColor seriesColor = OfficeColor.FromRgba(17, 34, 51, 128);
            OfficeColor pointColor = OfficeColor.FromRgba(68, 85, 102, 96);
            OfficeColor outlineColor = OfficeColor.FromRgba(119, 136, 153, 64);
            var data = new OfficeChartData(new[] { "1" }, new[] {
                OfficeChartSeries.CreateBubble(
                    "Portfolio",
                    new[] { 1D },
                    new[] { 2D },
                    new[] { 4D },
                    seriesColor,
                    new OfficeColor?[] { pointColor },
                    markerOutlineColor: outlineColor)
            });

            PowerPointChart chart = presentation.AddSlide().AddChart(
                OfficeChartKind.Bubble, data);
            C.BubbleChartSeries nativeSeries = presentation.Slides[0].SlidePart
                .ChartParts.Single().ChartSpace!
                .Descendants<C.BubbleChartSeries>().Single();
            C.ChartShapeProperties seriesProperties =
                nativeSeries.GetFirstChild<C.ChartShapeProperties>()!;
            Assert.Equal(50196, seriesProperties
                .GetFirstChild<DocumentFormat.OpenXml.Drawing.SolidFill>()!
                .RgbColorModelHex!.GetFirstChild<DocumentFormat.OpenXml.Drawing.Alpha>()!
                .Val!.Value);
            Assert.Equal(25098, seriesProperties
                .GetFirstChild<DocumentFormat.OpenXml.Drawing.Outline>()!
                .GetFirstChild<DocumentFormat.OpenXml.Drawing.SolidFill>()!
                .RgbColorModelHex!.GetFirstChild<DocumentFormat.OpenXml.Drawing.Alpha>()!
                .Val!.Value);
            Assert.Equal(37647, nativeSeries.GetFirstChild<C.DataPoint>()!
                .GetFirstChild<C.ChartShapeProperties>()!
                .GetFirstChild<DocumentFormat.OpenXml.Drawing.SolidFill>()!
                .RgbColorModelHex!.GetFirstChild<DocumentFormat.OpenXml.Drawing.Alpha>()!
                .Val!.Value);

            Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
            OfficeChartSeries snapshotSeries = Assert.Single(snapshot.Data.Series);
            Assert.Equal(seriesColor, snapshotSeries.Color);
            Assert.Equal(pointColor, Assert.Single(snapshotSeries.PointColors!));
            Assert.Equal(outlineColor, snapshotSeries.MarkerOutlineColor);
        }

        [Fact]
        public void BubbleChart_PreservesExplicitlyDisabledSeriesOutline() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create(new MemoryStream());
            PowerPointChart chart = presentation.AddSlide().AddChart(
                OfficeChartKind.Bubble,
                CreateBubbleData(
                    new[] { 1D },
                    new[] { 2D },
                    new[] { 4D },
                    seriesColor: OfficeColor.Parse("#112233")));
            C.BubbleChartSeries nativeSeries = presentation.Slides[0].SlidePart
                .ChartParts.Single().ChartSpace!
                .Descendants<C.BubbleChartSeries>().Single();
            DocumentFormat.OpenXml.Drawing.Outline outline = nativeSeries
                .GetFirstChild<C.ChartShapeProperties>()!
                .GetFirstChild<DocumentFormat.OpenXml.Drawing.Outline>()!;
            outline.RemoveAllChildren<DocumentFormat.OpenXml.Drawing.SolidFill>();
            outline.Append(new DocumentFormat.OpenXml.Drawing.NoFill());

            Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
            OfficeChartSeries series = Assert.Single(snapshot.Data.Series);
            Assert.False(series.ShowMarkerOutline);
            OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(snapshot);
            OfficeDrawingShape bubble = Assert.Single(drawing.Shapes, shape =>
                shape.Shape.Kind == OfficeShapeKind.Ellipse &&
                shape.Shape.FillColor == OfficeColor.Parse("#112233"));
            Assert.Null(bubble.Shape.StrokeColor);
            Assert.Equal(0D, bubble.Shape.StrokeWidth);
        }

        [Fact]
        public void BubbleChart_RejectsMixedSnapshotsThatWouldDropBubbleSeries() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create(new MemoryStream());
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointChart bubble = slide.AddChart(
                OfficeChartKind.Bubble,
                CreateBubbleData(
                    new[] { 1D, 2D },
                    new[] { 2D, 4D },
                    new[] { 4D, 9D }));
            OfficeChartData categoryData = new(
                new[] { "A", "B" },
                new[] { new OfficeChartSeries("Values", new[] { 1D, 2D }) });
            slide.AddChart(OfficeChartKind.ColumnClustered, categoryData);
            slide.AddChart(OfficeChartKind.Line, categoryData);

            ChartPart[] chartParts = slide.SlidePart.ChartParts.ToArray();
            C.PlotArea bubblePlot = chartParts[0].ChartSpace!
                .GetFirstChild<C.Chart>()!.GetFirstChild<C.PlotArea>()!;
            C.PlotArea columnPlot = chartParts[1].ChartSpace!
                .GetFirstChild<C.Chart>()!.GetFirstChild<C.PlotArea>()!;
            C.PlotArea linePlot = chartParts[2].ChartSpace!
                .GetFirstChild<C.Chart>()!.GetFirstChild<C.PlotArea>()!;
            bubblePlot.InsertBefore(
                columnPlot.GetFirstChild<C.BarChart>()!.CloneNode(true),
                bubblePlot.GetFirstChild<C.ValueAxis>());
            bubblePlot.InsertBefore(
                linePlot.GetFirstChild<C.LineChart>()!.CloneNode(true),
                bubblePlot.GetFirstChild<C.ValueAxis>());

            Assert.False(bubble.TryGetOfficeSnapshot(out _));
        }

        [Fact]
        public void BubbleChart_DrivesPngSvgHtmlAndPdfWithoutFallbackDiagnostics() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(480, 300);
            PowerPointSlide slide = presentation.AddSlide();
            slide.AddChartPoints(OfficeChartKind.Bubble,
                CreateBubbleData(
                    new[] { 1D, 2D, 3D },
                    new[] { 2D, 5D, 4D },
                    new[] { 4D, 25D, 9D },
                    new OfficeColor?[] {
                        OfficeColor.Parse("#D62828"),
                        OfficeColor.Parse("#2A9D8F"),
                        OfficeColor.Parse("#F4A261")
                    }),
                30, 20, 420, 250);

            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            string svgText = System.Text.Encoding.UTF8.GetString(svg.Bytes);
            Assert.True(png.Bytes.Length > 100);
            Assert.Contains("<ellipse", svgText, StringComparison.Ordinal);
            Assert.Contains("#D62828", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain(png.Diagnostics,
                diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.DoesNotContain(svg.Diagnostics,
                diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);

            PowerPointToHtmlResult htmlResult = presentation.ToHtmlResult(new PowerPointHtmlSaveOptions {
                Profile = OfficeHtmlConversionProfile.PowerPointVisualReview
            });
            string html = htmlResult.Value;
            PdfDocumentConversionResult pdfResult = presentation.ToPdfDocumentResult();
            byte[] pdf = pdfResult.ToBytes();
            Assert.Contains("<ellipse", html, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain(htmlResult.ImageDiagnostics,
                diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.DoesNotContain(pdfResult.Warnings,
                warning => warning.Code == "snapshot-selective-fallback");
            Assert.DoesNotContain(pdfResult.Warnings,
                warning => warning.Code == "unsupported-chart");
            Assert.True(pdf.Length > 500);
        }

        [Fact]
        public void BubbleChart_ZeroSizesRemainInvisibleInSharedRendering() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create(new MemoryStream());
            PowerPointSlide slide = presentation.AddSlide();
            slide.AddChartPoints(OfficeChartKind.Bubble,
                CreateBubbleData(
                    new[] { 1D, 2D },
                    new[] { 2D, 4D },
                    new[] { 0D, 9D }),
                30, 20, 420, 250);

            string svg = System.Text.Encoding.UTF8.GetString(
                slide.ExportImage(OfficeImageExportFormat.Svg).Bytes);

            Assert.Equal(1, CountOccurrences(svg, "<ellipse"));
        }

        [Fact]
        public void BubbleChart_RejectsMissingOrInvalidSizesBeforeMutatingSlide() {
            Assert.Throws<ArgumentException>(() => OfficeChartSeries.CreateBubble("Invalid",
                new[] { 1D, 2D }, new[] { 2D, 3D }, new[] { 5D }));
            Assert.Throws<ArgumentOutOfRangeException>(() => OfficeChartSeries.CreateBubble("Invalid",
                new[] { 1D }, new[] { 2D }, new[] { -1D }));
            Assert.Throws<ArgumentOutOfRangeException>(() => OfficeChartSeries.CreateBubble("Invalid",
                new[] { double.NaN }, new[] { 2D }, new[] { 1D }));
            Assert.Throws<ArgumentOutOfRangeException>(() => OfficeChartSeries.CreateBubble("Invalid",
                new[] { 1D }, new[] { double.PositiveInfinity }, new[] { 1D }));

            using PowerPointPresentation presentation =
                PowerPointPresentation.Create(new MemoryStream());
            PowerPointSlide slide = presentation.AddSlide();
            var missingSizes = new OfficeChartData(new[] { "1", "2" }, new[] {
                new OfficeChartSeries("Incomplete", new[] { 3D, 4D }, new[] { 1D, 2D })
            });
            Assert.Throws<ArgumentException>(() =>
                slide.AddChart(OfficeChartKind.Bubble, missingSizes));
            Assert.Empty(slide.Charts);

            PowerPointChart imported = slide.AddChart(OfficeChartKind.Bubble,
                CreateBubbleData(new[] { 1D }, new[] { 2D }, new[] { 4D }));
            C.BubbleSize bubbleSize = slide.SlidePart.ChartParts.Single().ChartSpace!
                .Descendants<C.BubbleSize>().Single();
            bubbleSize.NumberReference!.NumberingCache!.Elements<C.NumericPoint>()
                .Single().NumericValue!.Text = "-4";
            Assert.False(imported.TryGetOfficeSnapshot(out _));
        }

        [Fact]
        public void BubbleChart_RejectsMismatchedImportedCaches() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create(new MemoryStream());
            PowerPointChart chart = presentation.AddSlide().AddChart(
                OfficeChartKind.Bubble,
                CreateBubbleData(
                    new[] { 1D, 2D, 3D },
                    new[] { 2D, 4D, 6D },
                    new[] { 4D, 9D, 16D }));
            C.NumberingCache sizeCache = presentation.Slides[0].SlidePart.ChartParts.Single()
                .ChartSpace!.Descendants<C.BubbleSize>().Single()
                .NumberReference!.NumberingCache!;
            sizeCache.Elements<C.NumericPoint>().Last().Remove();
            sizeCache.PointCount!.Val = 2U;

            Assert.False(chart.TryGetOfficeSnapshot(out _));
        }

        [Fact]
        public void BubbleChart_RejectsMalformedImportedSizeCaches() {
            foreach (string malformed in new[] { "NaN", "Infinity", "not-a-number" }) {
                using PowerPointPresentation presentation =
                    PowerPointPresentation.Create(new MemoryStream());
                PowerPointChart chart = presentation.AddSlide().AddChart(
                    OfficeChartKind.Bubble,
                    CreateBubbleData(
                        new[] { 1D },
                        new[] { 2D },
                        new[] { 4D }));
                C.NumericPoint point = presentation.Slides[0].SlidePart.ChartParts
                    .Single().ChartSpace!.Descendants<C.BubbleSize>().Single()
                    .NumberReference!.NumberingCache!.Elements<C.NumericPoint>().Single();
                point.NumericValue!.Text = malformed;

                Assert.False(chart.TryGetOfficeSnapshot(out _));
            }
        }

        [Fact]
        public void BubbleChart_RejectsEnabledThreeDimensionalRendering() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create(new MemoryStream());
            PowerPointChart chart = presentation.AddSlide().AddChart(
                OfficeChartKind.Bubble,
                CreateBubbleData(
                    new[] { 1D },
                    new[] { 2D },
                    new[] { 4D }));
            C.BubbleChart nativeChart = presentation.Slides[0].SlidePart.ChartParts
                .Single().ChartSpace!.Descendants<C.BubbleChart>().Single();
            C.BubbleChartSeries nativeSeries =
                nativeChart.Elements<C.BubbleChartSeries>().Single();

            nativeChart.GetFirstChild<C.Bubble3D>()!.Val = true;
            Assert.False(chart.TryGetOfficeSnapshot(out _));

            nativeChart.GetFirstChild<C.Bubble3D>()!.Val = false;
            nativeSeries.GetFirstChild<C.Bubble3D>()!.Val = true;
            Assert.False(chart.TryGetOfficeSnapshot(out _));

            nativeSeries.GetFirstChild<C.Bubble3D>()!.Val = false;
            Assert.True(chart.TryGetOfficeSnapshot(out _));

            nativeSeries.GetFirstChild<C.Bubble3D>()!.Val = null;
            Assert.False(chart.TryGetOfficeSnapshot(out _));

            nativeSeries.GetFirstChild<C.Bubble3D>()!.Val = false;
            C.DataPoint point = nativeSeries.PrependChild(new C.DataPoint(
                new C.Index { Val = 0U },
                new C.Bubble3D()));
            Assert.False(chart.TryGetOfficeSnapshot(out _));

            point.GetFirstChild<C.Bubble3D>()!.Val = false;
            Assert.True(chart.TryGetOfficeSnapshot(out _));
        }

        [Fact]
        public void BubbleChart_RejectsUnsupportedFillStyles() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create(new MemoryStream());
            PowerPointChart chart = presentation.AddSlide().AddChart(
                OfficeChartKind.Bubble,
                CreateBubbleData(
                    new[] { 1D },
                    new[] { 2D },
                    new[] { 4D },
                    seriesColor: OfficeColor.Parse("#112233")));
            C.BubbleChartSeries nativeSeries = presentation.Slides[0].SlidePart
                .ChartParts.Single().ChartSpace!
                .Descendants<C.BubbleChartSeries>().Single();
            C.ChartShapeProperties seriesProperties =
                nativeSeries.GetFirstChild<C.ChartShapeProperties>() ??
                nativeSeries.AppendChild(new C.ChartShapeProperties());
            seriesProperties.RemoveAllChildren<DocumentFormat.OpenXml.Drawing.SolidFill>();
            seriesProperties.PrependChild(new DocumentFormat.OpenXml.Drawing.NoFill());
            Assert.False(chart.TryGetOfficeSnapshot(out _));

            seriesProperties.RemoveAllChildren<DocumentFormat.OpenXml.Drawing.NoFill>();
            seriesProperties.PrependChild(
                new DocumentFormat.OpenXml.Drawing.GradientFill());
            Assert.False(chart.TryGetOfficeSnapshot(out _));

            seriesProperties.RemoveAllChildren<DocumentFormat.OpenXml.Drawing.GradientFill>();
            seriesProperties.PrependChild(new DocumentFormat.OpenXml.Drawing.SolidFill(
                new DocumentFormat.OpenXml.Drawing.RgbColorModelHex {
                    Val = "112233"
                }));
            C.DataPoint point = nativeSeries.PrependChild(new C.DataPoint(
                new C.Index { Val = 0U },
                new C.ChartShapeProperties(
                    new DocumentFormat.OpenXml.Drawing.NoFill())));
            Assert.False(chart.TryGetOfficeSnapshot(out _));

            C.ChartShapeProperties pointProperties =
                point.GetFirstChild<C.ChartShapeProperties>()!;
            pointProperties.RemoveAllChildren<DocumentFormat.OpenXml.Drawing.NoFill>();
            pointProperties.Append(new DocumentFormat.OpenXml.Drawing.PatternFill());
            Assert.False(chart.TryGetOfficeSnapshot(out _));

            pointProperties.RemoveAllChildren<DocumentFormat.OpenXml.Drawing.PatternFill>();
            pointProperties.Append(new DocumentFormat.OpenXml.Drawing.Outline(
                new DocumentFormat.OpenXml.Drawing.SolidFill(
                    new DocumentFormat.OpenXml.Drawing.RgbColorModelHex {
                        Val = "445566"
                    })));
            Assert.False(chart.TryGetOfficeSnapshot(out _));
        }

        [Fact]
        public void BubbleChart_RejectsUnsupportedSeriesOutlineFills() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create(new MemoryStream());
            PowerPointChart chart = presentation.AddSlide().AddChart(
                OfficeChartKind.Bubble,
                CreateBubbleData(
                    new[] { 1D },
                    new[] { 2D },
                    new[] { 4D },
                    seriesColor: OfficeColor.Parse("#112233")));
            C.BubbleChartSeries nativeSeries = presentation.Slides[0].SlidePart
                .ChartParts.Single().ChartSpace!
                .Descendants<C.BubbleChartSeries>().Single();
            C.ChartShapeProperties properties =
                nativeSeries.GetFirstChild<C.ChartShapeProperties>() ??
                nativeSeries.AppendChild(new C.ChartShapeProperties());
            DocumentFormat.OpenXml.Drawing.Outline outline =
                properties.GetFirstChild<DocumentFormat.OpenXml.Drawing.Outline>() ??
                properties.AppendChild(new DocumentFormat.OpenXml.Drawing.Outline());
            outline.RemoveAllChildren<DocumentFormat.OpenXml.Drawing.SolidFill>();
            outline.Append(new DocumentFormat.OpenXml.Drawing.GradientFill());

            Assert.False(chart.TryGetOfficeSnapshot(out _));

            outline.RemoveAllChildren<DocumentFormat.OpenXml.Drawing.GradientFill>();
            outline.Append(new DocumentFormat.OpenXml.Drawing.PresetDash {
                Val = DocumentFormat.OpenXml.Drawing.PresetLineDashValues.Dash
            });
            Assert.False(chart.TryGetOfficeSnapshot(out _));

            outline.RemoveAllChildren<DocumentFormat.OpenXml.Drawing.PresetDash>();
            outline.Append(new DocumentFormat.OpenXml.Drawing.CustomDash());
            Assert.False(chart.TryGetOfficeSnapshot(out _));
        }

        [Fact]
        public void BubbleChart_RejectsNativeVaryColors() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create(new MemoryStream());
            PowerPointChart chart = presentation.AddSlide().AddChart(
                OfficeChartKind.Bubble,
                CreateBubbleData(
                    new[] { 1D },
                    new[] { 2D },
                    new[] { 4D }));
            C.VaryColors varyColors = presentation.Slides[0].SlidePart.ChartParts
                .Single().ChartSpace!.Descendants<C.BubbleChart>().Single()
                .GetFirstChild<C.VaryColors>()!;

            varyColors.Val = true;
            Assert.False(chart.TryGetOfficeSnapshot(out _));

            varyColors.Val = null;
            Assert.False(chart.TryGetOfficeSnapshot(out _));

            varyColors.Val = false;
            Assert.True(chart.TryGetOfficeSnapshot(out _));
        }

        [Fact]
        public void BubbleChart_RejectsTrendlinesAndErrorBarsFromSharedSnapshots() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create(new MemoryStream());
            PowerPointChart chart = presentation.AddSlide().AddChart(
                OfficeChartKind.Bubble,
                CreateBubbleData(
                    new[] { 1D, 2D },
                    new[] { 2D, 4D },
                    new[] { 4D, 9D }));

            chart.SetSeriesTrendline(0, C.TrendlineValues.Linear);

            Assert.NotNull(presentation.Slides[0].SlidePart.ChartParts.Single()
                .ChartSpace!.Descendants<C.Trendline>().Single());
            Assert.False(chart.TryGetOfficeSnapshot(out _));

            C.BubbleChartSeries series = presentation.Slides[0].SlidePart
                .ChartParts.Single().ChartSpace!
                .Descendants<C.BubbleChartSeries>().Single();
            series.RemoveAllChildren<C.Trendline>();
            series.Append(new C.ErrorBars());
            Assert.False(chart.TryGetOfficeSnapshot(out _));
        }

        private static int CountOccurrences(string value, string marker) {
            int count = 0;
            int offset = 0;
            while ((offset = value.IndexOf(marker, offset, StringComparison.Ordinal)) >= 0) {
                count++;
                offset += marker.Length;
            }
            return count;
        }

        private static OfficeChartData CreateBubbleData(IReadOnlyList<double> xValues,
            IReadOnlyList<double> yValues, IReadOnlyList<double> sizes,
            IReadOnlyList<OfficeColor?>? pointColors = null, OfficeColor? seriesColor = null) =>
            new(xValues.Select(value => value.ToString(System.Globalization.CultureInfo.InvariantCulture)),
                new[] {
                    OfficeChartSeries.CreateBubble("Portfolio", xValues, yValues, sizes,
                        seriesColor, pointColors)
                });
    }
}
