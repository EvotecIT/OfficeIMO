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
