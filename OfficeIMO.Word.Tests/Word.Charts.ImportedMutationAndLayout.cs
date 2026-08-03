using System.Linq;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Word;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using WP14 = DocumentFormat.OpenXml.Office2010.Word.Drawing;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void ImportedChart_MutatesLiteralDataTitleAndFrameAcrossReload() {
            string filePath = System.IO.Path.Combine(_directoryWithFiles, "Chart.ImportedMutation.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                WordChart chart = document.AddChart("Original");
                chart.AddCategories(new List<string> { "A", "B", "C" });
                chart.AddBar("Before", new List<int> { 1, 2, 3 }, OfficeIMO.Drawing.OfficeColor.Blue);
                C.Values values = document._wordprocessingDocument.MainDocumentPart!.ChartParts.Single()
                    .ChartSpace.GetFirstChild<C.Chart>()!.PlotArea!.GetFirstChild<C.BarChart>()!
                    .GetFirstChild<C.BarChartSeries>()!.GetFirstChild<C.Values>()!;
                C.NumberLiteral originalLiteral = values.GetFirstChild<C.NumberLiteral>()!;
                var cache = new C.NumberingCache(
                    new C.FormatCode { Text = "$#,##0.00" },
                    (C.PointCount)originalLiteral.PointCount!.CloneNode(true));
                foreach (C.NumericPoint point in originalLiteral.Elements<C.NumericPoint>()) {
                    cache.Append((C.NumericPoint)point.CloneNode(true));
                }
                values.RemoveAllChildren();
                values.Append(new C.NumberReference(
                    new C.Formula { Text = "Sheet1!$B$2:$B$4" },
                    cache));
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordChart imported = Assert.Single(document.Charts);

                imported.SetTitle("Imported mutation");
                Assert.True(imported.TrySetCategories(new[] { "North", "South", "West" }));
                Assert.True(imported.TrySetSeriesName(0, "After"));
                Assert.True(imported.TrySetSeriesValues(0, new[] { 10.5, 20.25, 30.75 }));
                imported.SetSize(320, 180);

                Assert.True(imported.TryGetSnapshot(out WordChartSnapshot snapshot));
                Assert.Equal("Imported mutation", snapshot.Title);
                Assert.Equal(new[] { "North", "South", "West" }, snapshot.Data.Categories);
                WordChartSeries series = Assert.Single(snapshot.Data.Series);
                Assert.Equal("After", series.Name);
                Assert.Equal(new[] { 10.5, 20.25, 30.75 }, series.Values);
                Assert.Equal(240D, snapshot.WidthPoints, 6);
                Assert.Equal(135D, snapshot.HeightPoints, 6);
                document.Save();
            }

            using WordDocument reloaded = WordDocument.Load(filePath);
            WordChart persisted = Assert.Single(reloaded.Charts);
            Assert.True(persisted.TryGetSnapshot(out WordChartSnapshot persistedSnapshot));
            Assert.Equal("Imported mutation", persistedSnapshot.Title);
            Assert.Equal(new[] { "North", "South", "West" }, persistedSnapshot.Data.Categories);
            Assert.Equal(new[] { 10.5, 20.25, 30.75 }, Assert.Single(persistedSnapshot.Data.Series).Values);

            C.BarChartSeries xmlSeries = reloaded._wordprocessingDocument.MainDocumentPart!.ChartParts
                .Single().ChartSpace.GetFirstChild<C.Chart>()!.PlotArea!.GetFirstChild<C.BarChart>()!
                .GetFirstChild<C.BarChartSeries>()!;
            Assert.NotNull(xmlSeries.GetFirstChild<C.CategoryAxisData>()!.GetFirstChild<C.StringLiteral>());
            Assert.Null(xmlSeries.GetFirstChild<C.CategoryAxisData>()!.GetFirstChild<C.StringReference>());
            C.NumberLiteral persistedLiteral = Assert.IsType<C.NumberLiteral>(xmlSeries.GetFirstChild<C.Values>()!.GetFirstChild<C.NumberLiteral>());
            Assert.Equal("$#,##0.00", persistedLiteral.FormatCode!.Text);
            Assert.Null(xmlSeries.GetFirstChild<C.Values>()!.GetFirstChild<C.NumberReference>());
            Assert.Empty(new OpenXmlValidator().Validate(reloaded._wordprocessingDocument));
        }

        [Fact]
        public void AnchoredChart_IsDiscoveredSizedAndReportedAsPackageGeometry() {
            string filePath = System.IO.Path.Combine(_directoryWithFiles, "Chart.AnchoredLayout.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                WordChart chart = document.AddChart("Anchored");
                chart.AddCategories(new List<string> { "A", "B" });
                chart.AddBar("Series", new List<int> { 1, 2 }, OfficeIMO.Drawing.OfficeColor.Green);
                ConvertChartToPageAnchor(chart, 36, 72);
                DW.Anchor anchor = chart.Drawing!.Anchor!;
                anchor.Append(new WP14.RelativeWidth(
                    new WP14.PercentageWidth("50000")) {
                    ObjectId = WP14.SizeRelativeHorizontallyValues.Page
                });
                anchor.Append(new WP14.RelativeHeight(
                    new WP14.PercentageHeight("50000")) {
                    RelativeFrom = WP14.SizeRelativeVerticallyValues.Page
                });

                Assert.True(chart.TryGetLayoutSnapshot(out WordDrawingLayoutSnapshot layout));
                Assert.Equal(WordDrawingPlacementKind.Anchored, layout.Placement);
                Assert.Equal("page", layout.HorizontalRelativeFrom);
                Assert.Equal("page", layout.VerticalRelativeFrom);
                Assert.Equal(36D, layout.HorizontalOffsetPoints!.Value, 6);
                Assert.Equal(72D, layout.VerticalOffsetPoints!.Value, 6);
                Assert.Equal(WordDrawingWrapKind.Square, layout.Wrap);
                chart.SetSize(400, 200);
                Assert.Null(anchor.GetFirstChild<WP14.RelativeWidth>());
                Assert.Null(anchor.GetFirstChild<WP14.RelativeHeight>());
                document.Save();
            }

            using WordDocument reloaded = WordDocument.Load(filePath);
            WordChart imported = Assert.Single(reloaded.Charts);
            Assert.True(imported.TryGetLayoutSnapshot(out WordDrawingLayoutSnapshot persisted));
            Assert.Equal(WordDrawingPlacementKind.Anchored, persisted.Placement);
            Assert.Equal(300D, persisted.WidthPoints, 6);
            Assert.Equal(150D, persisted.HeightPoints, 6);
            Assert.True(imported.TryGetSnapshot(out WordChartSnapshot chartSnapshot));
            Assert.Equal(300D, chartSnapshot.WidthPoints, 6);
            Assert.Empty(new OpenXmlValidator().Validate(reloaded._wordprocessingDocument));
        }

        [Fact]
        public void ShapeAndSmartArt_UseTheSharedPersistedLayoutEvidence() {
            using WordDocument document = WordDocument.Create();
            WordShape shape = WordShape.AddDrawingShapeAnchored(
                document.AddParagraph("shape"),
                ShapeType.Rectangle,
                80,
                40,
                18,
                27);
            WordSmartArt smartArt = document.AddSmartArt(SmartArtType.BasicProcess);

            Assert.True(shape.TryGetLayoutSnapshot(out WordDrawingLayoutSnapshot shapeLayout));
            Assert.Equal(WordDrawingPlacementKind.Anchored, shapeLayout.Placement);
            Assert.Equal(18D, shapeLayout.HorizontalOffsetPoints!.Value, 6);
            Assert.Equal(27D, shapeLayout.VerticalOffsetPoints!.Value, 6);
            Assert.Equal(WordDrawingWrapKind.Square, shapeLayout.Wrap);

            Assert.True(smartArt.TryGetLayoutSnapshot(out WordDrawingLayoutSnapshot smartArtLayout));
            Assert.Equal(WordDrawingPlacementKind.Inline, smartArtLayout.Placement);
            Assert.Equal(432D, smartArtLayout.WidthPoints, 6);
            Assert.Equal(252D, smartArtLayout.HeightPoints, 6);
        }

        [Fact]
        public void ImportedSimplePositionAnchor_UsesPageCoordinatesInsteadOfPositionPlaceholders() {
            string filePath = System.IO.Path.Combine(_directoryWithFiles, "Chart.SimplePositionLayout.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                WordChart chart = document.AddChart("Simple position");
                chart.AddCategories(new List<string> { "A" });
                chart.AddBar("Series", new List<int> { 1 }, OfficeIMO.Drawing.OfficeColor.Blue);
                ConvertChartToPageAnchor(chart, 300, 400);
                DW.Anchor anchor = chart.Drawing!.Anchor!;
                anchor.SimplePos = true;
                anchor.SimplePosition!.X = (long)(21D * 12700D);
                anchor.SimplePosition.Y = (long)(33D * 12700D);

                Assert.True(chart.TryGetLayoutSnapshot(out WordDrawingLayoutSnapshot layout));
                Assert.True(layout.UsesSimplePosition);
                Assert.Equal("page", layout.HorizontalRelativeFrom);
                Assert.Equal("page", layout.VerticalRelativeFrom);
                Assert.Equal(21D, layout.HorizontalOffsetPoints!.Value, 6);
                Assert.Equal(33D, layout.VerticalOffsetPoints!.Value, 6);
                Assert.Null(layout.HorizontalAlignment);
                Assert.Null(layout.VerticalAlignment);
                document.Save();
            }

            using WordDocument reloaded = WordDocument.Load(filePath);
            WordChart imported = Assert.Single(reloaded.Charts);
            Assert.True(imported.TryGetLayoutSnapshot(out WordDrawingLayoutSnapshot persisted));
            Assert.True(persisted.UsesSimplePosition);
            Assert.Equal(21D, persisted.HorizontalOffsetPoints!.Value, 6);
            Assert.Equal(33D, persisted.VerticalOffsetPoints!.Value, 6);
            Assert.Empty(new OpenXmlValidator().Validate(reloaded._wordprocessingDocument));
        }

        private static void ConvertChartToPageAnchor(WordChart chart, double leftPoints, double topPoints) {
            DocumentFormat.OpenXml.Wordprocessing.Drawing drawing = chart.Drawing!;
            Inline inline = drawing.Inline!;
            var anchor = new DW.Anchor {
                DistanceFromTop = 0U,
                DistanceFromBottom = 0U,
                DistanceFromLeft = 0U,
                DistanceFromRight = 0U,
                SimplePos = false,
                RelativeHeight = 5U,
                BehindDoc = false,
                Locked = false,
                LayoutInCell = true,
                AllowOverlap = true
            };
            anchor.Append(new SimplePosition { X = 0L, Y = 0L });
            anchor.Append(new HorizontalPosition(
                new PositionOffset(((long)(leftPoints * 12700D)).ToString(System.Globalization.CultureInfo.InvariantCulture))) {
                RelativeFrom = HorizontalRelativePositionValues.Page
            });
            anchor.Append(new VerticalPosition(
                new PositionOffset(((long)(topPoints * 12700D)).ToString(System.Globalization.CultureInfo.InvariantCulture))) {
                RelativeFrom = VerticalRelativePositionValues.Page
            });
            anchor.Append((Extent)inline.Extent!.CloneNode(true));
            anchor.Append((EffectExtent)inline.EffectExtent!.CloneNode(true));
            anchor.Append(new WrapSquare { WrapText = WrapTextValues.BothSides });
            anchor.Append((DocProperties)inline.DocProperties!.CloneNode(true));
            anchor.Append(inline.NonVisualGraphicFrameDrawingProperties != null
                ? (DW.NonVisualGraphicFrameDrawingProperties)inline.NonVisualGraphicFrameDrawingProperties.CloneNode(true)
                : new DW.NonVisualGraphicFrameDrawingProperties(new GraphicFrameLocks { NoChangeAspect = true }));
            anchor.Append((Graphic)inline.Graphic!.CloneNode(true));
            drawing.RemoveAllChildren();
            drawing.Append(anchor);
        }
    }
}
