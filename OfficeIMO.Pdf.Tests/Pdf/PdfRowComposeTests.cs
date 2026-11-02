using System;
using System.Linq;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf {
    public class PdfRowBuilderTests {
        [Fact]
        public void RowRejectsEmptyComposition() {
            var doc = PdfDocument.Create();

            var exception = Assert.Throws<InvalidOperationException>(() =>
                doc.Compose(compose =>
                    compose.Page(page =>
                        page.Content(content =>
                            content.Row(_ => { })))));

            Assert.Contains("Rows require at least one column.", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void ColumnRejectsNonPositiveWidth() {
            var doc = PdfDocument.Create();

            Assert.Throws<ArgumentOutOfRangeException>(() =>
                doc.Compose(compose =>
                    compose.Page(page =>
                        page.Content(content =>
                            content.Row(row =>
                                row.PercentColumn(0, _ => { }))))));
        }

        [Theory]
        [InlineData(double.NaN)]
        [InlineData(double.PositiveInfinity)]
        [InlineData(double.NegativeInfinity)]
        public void ColumnRejectsNonFiniteWidth(double widthPercent) {
            var doc = PdfDocument.Create();

            var exception = Assert.Throws<ArgumentOutOfRangeException>(() =>
                doc.Compose(compose =>
                    compose.Page(page =>
                        page.Content(content =>
                            content.Row(row =>
                                row.PercentColumn(widthPercent, _ => { }))))));

            Assert.Equal("percent", exception.ParamName);
            Assert.Contains("Column width must be finite and greater than zero.", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void ColumnRejectsWidthOverOneHundred() {
            var doc = PdfDocument.Create();

            Assert.Throws<ArgumentOutOfRangeException>(() =>
                doc.Compose(compose =>
                    compose.Page(page =>
                        page.Content(content =>
                            content.Row(row =>
                                row.PercentColumn(150, _ => { }))))));
        }

        [Fact]
        public void ColumnRejectsWhenTotalWouldExceedOneHundred() {
            var doc = PdfDocument.Create();

            Assert.Throws<InvalidOperationException>(() =>
                doc.Compose(compose =>
                    compose.Page(page =>
                        page.Content(content =>
                            content.Row(row => {
                                row.PercentColumn(60, _ => { });
                                row.PercentColumn(50, _ => { });
                            })))));
        }

        [Fact]
        public void PercentageColumnIntentIsPreservedWhenTotalLessThanOneHundred() {
            var doc = PdfDocument.Create();

            doc.Compose(compose =>
                compose.Page(page =>
                    page.Content(content =>
                        content.Row(row => {
                            row.PercentColumn(30, _ => { });
                            row.PercentColumn(20, _ => { });
                        }))));

            var page = Assert.IsType<PageBlock>(Assert.Single(doc.Blocks));
            var row = Assert.IsType<RowBlock>(Assert.Single(page.Blocks));
            Assert.Equal(2, row.Columns.Count);

            Assert.All(row.Columns, column => Assert.Equal(PdfColumnWidthUnit.Percent, column.Width.Unit));
            Assert.Equal(30D, row.Columns[0].Width.Value);
            Assert.Equal(20D, row.Columns[1].Width.Value);

            string text = doc.Read().Text;
            Assert.NotNull(text);
        }

        [Fact]
        public void RowGap_IsStoredOnModel() {
            var doc = PdfDocument.Create();

            doc.Compose(compose =>
                compose.Page(page =>
                    page.Content(content =>
                        content.Row(row => {
                            row.Gap(global::OfficeIMO.OfficeLength.Points(18));
                            row.PercentColumn(50, _ => { });
                            row.PercentColumn(50, _ => { });
                        }))));

            var page = Assert.IsType<PageBlock>(Assert.Single(doc.Blocks));
            var row = Assert.IsType<RowBlock>(Assert.Single(page.Blocks));
            Assert.Equal(18, row.Gap);
        }

        [Fact]
        public void RowStyle_IsStoredOnModelAndSnapshotsInput() {
            var doc = PdfDocument.Create();
            var style = new PdfRowStyle {
                Gap = global::OfficeIMO.OfficeLength.Points(18),
                SpacingBefore = global::OfficeIMO.OfficeLength.Points(7),
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(9),
                ColumnSeparatorColor = global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.12, 0.34, 0.56),
                ColumnSeparatorWidth = global::OfficeIMO.OfficeLength.Points(1.25),
                KeepTogether = true,
                KeepWithNext = true
            };

            doc.Compose(compose =>
                compose.Page(page =>
                    page.Content(content =>
                        content.Row(row => {
                            row.Style(style);
                            row.PercentColumn(50, column => column.Paragraph(paragraph => paragraph.Text("Left")));
                            row.PercentColumn(50, column => column.Paragraph(paragraph => paragraph.Text("Right")));
                        }))));

            style.Gap = global::OfficeIMO.OfficeLength.Points(4);
            style.SpacingBefore = global::OfficeIMO.OfficeLength.Points(1);
            style.SpacingAfter = global::OfficeIMO.OfficeLength.Points(2);
            style.ColumnSeparatorColor = global::OfficeIMO.OfficeColor.Black;
            style.ColumnSeparatorWidth = global::OfficeIMO.OfficeLength.Points(0.5);
            style.KeepTogether = false;
            style.KeepWithNext = false;

            var page = Assert.IsType<PageBlock>(Assert.Single(doc.Blocks));
            var row = Assert.IsType<RowBlock>(Assert.Single(page.Blocks));

            Assert.Equal(18, row.Gap);
            Assert.Equal(18, row.Style!.Gap?.ToPoints());
            Assert.Equal(7, row.Style.SpacingBefore.ToPoints());
            Assert.Equal(9, row.Style.SpacingAfter.ToPoints());
            Assert.Equal(global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.12, 0.34, 0.56), row.Style.ColumnSeparatorColor);
            Assert.Equal(1.25, row.Style.ColumnSeparatorWidth.ToPoints());
            Assert.True(row.Style.KeepTogether);
            Assert.True(row.Style.KeepWithNext);
        }

        [Fact]
        public void RowUsesBuiltInWordLikeGapUnlessExplicitlyOptedOut() {
            var defaultDoc = PdfDocument.Create();
            defaultDoc.Compose(compose =>
                compose.Page(page =>
                    page.Content(content =>
                        content.Row(row => {
                            row.PercentColumn(50, column => column.Paragraph(paragraph => paragraph.Text("Left")));
                            row.PercentColumn(50, column => column.Paragraph(paragraph => paragraph.Text("Right")));
                        }))));

            var defaultPage = Assert.IsType<PageBlock>(Assert.Single(defaultDoc.Blocks));
            var defaultRow = Assert.IsType<RowBlock>(Assert.Single(defaultPage.Blocks));

            Assert.Equal(PdfRowStyle.DefaultGap.ToPoints(), defaultRow.Gap);

            var optOutDoc = PdfDocument.Create();
            optOutDoc.Compose(compose =>
                compose.Page(page =>
                    page.Content(content =>
                        content.Row(row => {
                            row.Gap(global::OfficeIMO.OfficeLength.Points(0));
                            row.PercentColumn(50, column => column.Paragraph(paragraph => paragraph.Text("Left")));
                            row.PercentColumn(50, column => column.Paragraph(paragraph => paragraph.Text("Right")));
                        }))));

            var optOutPage = Assert.IsType<PageBlock>(Assert.Single(optOutDoc.Blocks));
            var optOutRow = Assert.IsType<RowBlock>(Assert.Single(optOutPage.Blocks));

            Assert.Equal(0, optOutRow.Gap);
        }

        [Theory]
        [InlineData(-1)]



        public void RowGap_RejectsInvalidValues(double gap) {
            var doc = PdfDocument.Create();

            var exception = Assert.Throws<ArgumentOutOfRangeException>(() =>
                doc.Compose(compose =>
                    compose.Page(page =>
                        page.Content(content =>
                            content.Row(row => row.Gap(global::OfficeIMO.OfficeLength.Points(gap)))))));

            Assert.Equal("gap", exception.ParamName);
        }

        [Fact]
        public void RowStyle_RejectsInvalidValues() {

            Assert.Throws<ArgumentException>(() => new PdfRowStyle { Gap = global::OfficeIMO.OfficeLength.Points(-1) });

            Assert.Throws<ArgumentException>(() => new PdfRowStyle { SpacingAfter = global::OfficeIMO.OfficeLength.Points(-1) });

            Assert.Throws<ArgumentException>(() => new PdfRowStyle { ColumnSeparatorWidth = global::OfficeIMO.OfficeLength.Points(-1) });

            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Compose(compose =>
                    compose.Page(page =>
                        page.Content(content =>
                            content.Row(row => row.Style(null!))))));
        }

        [Fact]
        public void RowColumnSeparator_RendersBetweenColumns() {
            byte[] bytes = PdfDocument.Create(new PdfCreateOptions {
                    PageWidth = global::OfficeIMO.OfficeLength.Points(360),
                    PageHeight = global::OfficeIMO.OfficeLength.Points(180),
                    MarginLeft = global::OfficeIMO.OfficeLength.Points(30),
                    MarginRight = global::OfficeIMO.OfficeLength.Points(30),
                    MarginTop = global::OfficeIMO.OfficeLength.Points(30),
                    MarginBottom = global::OfficeIMO.OfficeLength.Points(30),
                    DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10)
                })
                .Compose(compose =>
                    compose.Page(page =>
                        page.Content(content =>
                            content.Row(row => row
                                .Gap(global::OfficeIMO.OfficeLength.Points(20))
                                .ColumnSeparator(global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.12, 0.34, 0.56), global::OfficeIMO.OfficeLength.Points(1.25))
                                .PercentColumn(50, column => column.Paragraph(paragraph => paragraph.Text("LeftSeparatorMarker")))
                                .PercentColumn(50, column => column.Paragraph(paragraph => paragraph.Text("RightSeparatorMarker")))))))
                .ToBytes();

            string rawPdf = PdfTestContentText.ReadDecoded(bytes);

            Assert.Contains("0.12 0.34 0.56 RG", rawPdf, StringComparison.Ordinal);
            Assert.Contains("1.25 w", rawPdf, StringComparison.Ordinal);
            Assert.Contains("180 150 m 180 ", rawPdf, StringComparison.Ordinal);
        }

        [Fact]
        public void RowGap_RejectsWhenGapsExceedContentWidthDuringRender() {
            var doc = PdfDocument.Create(new PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(120),
                PageHeight = global::OfficeIMO.OfficeLength.Points(160),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(20),
                MarginRight = global::OfficeIMO.OfficeLength.Points(20),
                MarginTop = global::OfficeIMO.OfficeLength.Points(20),
                MarginBottom = global::OfficeIMO.OfficeLength.Points(20)
            }).Compose(compose =>
                compose.Page(page =>
                    page.Content(content =>
                        content.Row(row => {
                            row.Gap(global::OfficeIMO.OfficeLength.Points(90));
                            row.PercentColumn(50, column => column.Paragraph(paragraph => paragraph.Text("Left")));
                            row.PercentColumn(50, column => column.Paragraph(paragraph => paragraph.Text("Right")));
                        }))));

            var exception = Assert.Throws<ArgumentException>(() => doc.ToBytes());
            Assert.Contains("Row column gaps must be smaller than the available page content width.", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void RowModel_ExposesReadOnlyColumnAndBlockCollections() {
            var doc = PdfDocument.Create();

            doc.Compose(compose =>
                compose.Page(page =>
                    page.Content(content =>
                        content.Row(row => {
                            row.PercentColumn(40, column => column
                                .H1("Left")
                                .Paragraph(paragraph => paragraph.Text("Body")));
                            row.PercentColumn(60, column => column.H2("Right"));
                        }))));

            var page = Assert.IsType<PageBlock>(Assert.Single(doc.Blocks));
            var row = Assert.IsType<RowBlock>(Assert.Single(page.Blocks));

            Assert.False(row.Columns is System.Collections.Generic.List<RowColumn>);
            Assert.False(row.Columns[0].Blocks is System.Collections.Generic.List<IPdfBlock>);
            Assert.Equal(2, row.Columns.Count);
            Assert.Equal(2, row.Columns[0].Blocks.Count);
            Assert.IsType<HeadingBlock>(row.Columns[0].Blocks[0]);
            Assert.IsType<RichParagraphBlock>(row.Columns[0].Blocks[1]);
        }

        [Fact]
        public void MixedColumnSizing_RendersFixedAutoAndRelativeColumns() {
            byte[] bytes = PdfDocument.Create(new PdfCreateOptions {
                    PageWidth = global::OfficeIMO.OfficeLength.Points(400),
                    PageHeight = global::OfficeIMO.OfficeLength.Points(180),
                    MarginLeft = global::OfficeIMO.OfficeLength.Points(20),
                    MarginRight = global::OfficeIMO.OfficeLength.Points(20),
                    MarginTop = global::OfficeIMO.OfficeLength.Points(20),
                    MarginBottom = global::OfficeIMO.OfficeLength.Points(20),
                    DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10)
                })
                .Compose(document => document.Content(content => content.Row(row => row
                    .Gap(global::OfficeIMO.OfficeLength.Points(10))
                    .FixedColumn(global::OfficeIMO.OfficeLength.Points(80), column => column.Text("FixedMarker"))
                    .AutoColumn(column => column.Text("ID"), maximum: global::OfficeIMO.OfficeLength.Points(60))
                    .RelativeColumn(column => column.Text("RelativeMarker")))))
                .ToBytes();

            using var parsed = UglyToad.PdfPig.PdfDocument.Open(bytes);
            var letters = parsed.GetPage(1).Letters;
            double fixedX = letters.First(letter => letter.Value == "F").StartBaseLine.X;
            double autoX = letters.First(letter => letter.Value == "I").StartBaseLine.X;
            double relativeX = letters.First(letter => letter.Value == "R").StartBaseLine.X;

            Assert.InRange(fixedX, 19.9D, 20.1D);
            Assert.InRange(autoX, 109.9D, 110.1D);
            Assert.True(relativeX > autoX + 10D, "Expected the relative column after the measured automatic column.");
            Assert.True(relativeX < 180D, "Expected the automatic column to use its content width rather than the remaining row width.");
        }

        [Fact]
        public void MixedColumnSizing_PreservesEverySizingIntentInTheModel() {
            var doc = PdfDocument.Create(document => document.Content(content => content.Row(row => row
                .FixedColumn(global::OfficeIMO.OfficeLength.Points(72), column => column.Text("Fixed"))
                .AutoColumn(column => column.Text("Auto"), minimum: global::OfficeIMO.OfficeLength.Points(24), maximum: global::OfficeIMO.OfficeLength.Points(96))
                .RelativeColumn(column => column.Text("Primary"), weight: 3)
                .PercentColumn(20, column => column.Text("Percent")))));

            var row = Assert.IsType<RowBlock>(Assert.Single(doc.Blocks));

            Assert.Equal(PdfColumnWidthUnit.Points, row.Columns[0].Width.Unit);
            Assert.Equal(72D, row.Columns[0].Width.Value);
            Assert.Equal(PdfColumnWidthUnit.Auto, row.Columns[1].Width.Unit);
            Assert.Equal(24D, row.Columns[1].Width.Minimum.ToPoints());
            Assert.Equal(96D, row.Columns[1].Width.Maximum?.ToPoints());
            Assert.Equal(PdfColumnWidthUnit.Relative, row.Columns[2].Width.Unit);
            Assert.Equal(3D, row.Columns[2].Width.Value);
            Assert.Equal(PdfColumnWidthUnit.Percent, row.Columns[3].Width.Unit);
            Assert.Equal(20D, row.Columns[3].Width.Value);
        }

        [Theory]
        [InlineData(0)]
        [InlineData(-1)]
        [InlineData(double.NaN)]
        [InlineData(double.PositiveInfinity)]
        public void RelativeWidth_RejectsInvalidWeight(double weight) {
            Assert.Throws<ArgumentOutOfRangeException>(() => PdfColumnWidth.Relative(weight));
        }

        [Theory]
        [InlineData(0)]
        [InlineData(-1)]


        public void FixedWidth_RejectsInvalidPoints(double points) {
            Assert.Throws<ArgumentOutOfRangeException>(() => PdfColumnWidth.Fixed(global::OfficeIMO.OfficeLength.Points(points)));
        }

        [Fact]
        public void AutomaticWidth_RejectsInvalidBounds() {
            Assert.Throws<ArgumentOutOfRangeException>(() => PdfColumnWidth.Auto(global::OfficeIMO.OfficeLength.Points(-1)));


            Assert.Throws<ArgumentOutOfRangeException>(() => PdfColumnWidth.Auto(minimum: global::OfficeIMO.OfficeLength.Points(40), maximum: global::OfficeIMO.OfficeLength.Points(20)));
        }

        [Fact]
        public void CommittedColumnWidths_RejectOverflowDuringLayout() {
            PdfDocument doc = PdfDocument.Create(new PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(200),
                PageHeight = global::OfficeIMO.OfficeLength.Points(160),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(20),
                MarginRight = global::OfficeIMO.OfficeLength.Points(20),
                MarginTop = global::OfficeIMO.OfficeLength.Points(20),
                MarginBottom = global::OfficeIMO.OfficeLength.Points(20)
            }).Compose(document => document.Content(content => content.Row(row => row
                .Gap(global::OfficeIMO.OfficeLength.Points(10))
                .FixedColumn(global::OfficeIMO.OfficeLength.Points(100), column => column.Text("Left"))
                .FixedColumn(global::OfficeIMO.OfficeLength.Points(80), column => column.Text("Right")))));

            var exception = Assert.Throws<ArgumentException>(() => doc.ToBytes());
            Assert.Contains("exceed the available row width", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void RelativeColumnRejectsAZeroWidthRemainder() {
            PdfDocument doc = PdfDocument.Create(new PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(200),
                PageHeight = global::OfficeIMO.OfficeLength.Points(160),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(20),
                MarginRight = global::OfficeIMO.OfficeLength.Points(20),
                MarginTop = global::OfficeIMO.OfficeLength.Points(20),
                MarginBottom = global::OfficeIMO.OfficeLength.Points(20)
            }).Compose(document => document.Content(content => content.Row(row => row
                .Gap(global::OfficeIMO.OfficeLength.Points(0))
                .FixedColumn(global::OfficeIMO.OfficeLength.Points(160), column => column.Text("Fixed"))
                .RelativeColumn(column => column.Text("Relative")))));

            var exception = Assert.Throws<ArgumentException>(() => doc.ToBytes());
            Assert.Contains("require positive width", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void AutomaticColumnMeasuresPanelContentAndPadding() {
            byte[] bytes = PdfDocument.Create(new PdfCreateOptions {
                    PageWidth = global::OfficeIMO.OfficeLength.Points(400),
                    PageHeight = global::OfficeIMO.OfficeLength.Points(180),
                    MarginLeft = global::OfficeIMO.OfficeLength.Points(20),
                    MarginRight = global::OfficeIMO.OfficeLength.Points(20),
                    MarginTop = global::OfficeIMO.OfficeLength.Points(20),
                    MarginBottom = global::OfficeIMO.OfficeLength.Points(20),
                    DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10)
                })
                .Compose(document => document.Content(content => content.Row(row => row
                    .Gap(global::OfficeIMO.OfficeLength.Points(10))
                    .AutoColumn(column => column.PanelParagraph(
                        paragraph => paragraph.Text("LongPanelMarker"),
                        new PdfPanelStyle { PaddingX = global::OfficeIMO.OfficeLength.Points(18), PaddingY = global::OfficeIMO.OfficeLength.Points(2), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) }))
                    .RelativeColumn(column => column.Text("RelativeMarker")))))
                .ToBytes();

            using var parsed = UglyToad.PdfPig.PdfDocument.Open(bytes);
            var letters = parsed.GetPage(1).Letters;
            double panelX = letters.First(letter => letter.Value == "L").StartBaseLine.X;
            double relativeX = letters.First(letter => letter.Value == "R").StartBaseLine.X;

            Assert.True(relativeX > panelX + 95D, "Expected automatic sizing to include the panel text and horizontal padding.");
        }

        [Fact]
        public void RowColumn_UsesUniversalContentReceiverForAnnotations() {
            byte[] bytes = PdfDocument.Create(document => document.Content(content => content.Row(row => row
                    .RelativeColumn(column => column
                        .Text("Column content")
                        .TextAnnotation("Column note")))))
                .ToBytes();

            PdfDocument document = PdfDocument.Load(bytes);
            PdfAnnotation annotation = Assert.Single(document.Reader.AnnotationsBySubtype("Text"));

            Assert.Equal("Column note", annotation.Contents);
        }

        [Fact]
        public void RowColumn_ItemProvidesWordLikeFlowGroups() {
            var doc = PdfDocument.Create();

            doc.Compose(compose =>
                compose.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.PercentColumn(100, column => column
                                .Item(item => item
                                    .H2("Grouped heading")
                                    .Paragraph(paragraph => paragraph.Text("Grouped paragraph")))
                                .Spacer(global::OfficeIMO.OfficeLength.Points(6))
                                .Item(item => item.Bullets(new[] { "Grouped bullet" })))))));

            var page = Assert.IsType<PageBlock>(Assert.Single(doc.Blocks));
            var row = Assert.IsType<RowBlock>(Assert.Single(page.Blocks));
            var blocks = row.Columns[0].Blocks;

            Assert.Equal(4, blocks.Count);
            Assert.IsType<HeadingBlock>(blocks[0]);
            Assert.IsType<RichParagraphBlock>(blocks[1]);
            Assert.IsType<SpacerBlock>(blocks[2]);
            Assert.IsType<BulletListBlock>(blocks[3]);
        }

        [Fact]
        public void RowColumn_RejectsNullFlowDelegates() {
            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Compose(compose =>
                    compose.Page(page =>
                        page.Content(content =>
                            content.Row(row => row.PercentColumn(100, column => column.Item(null!)))))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Compose(compose =>
                    compose.Page(page =>
                        page.Content(content =>
                            content.Row(row => row.PercentColumn(100, column => column.Paragraph(null!)))))));
        }

        [Fact]
        public void RowColumn_CanComposeBulletAndNumberedLists() {
            var doc = PdfDocument.Create();
            var style = new PdfListStyle {
                FontSize = global::OfficeIMO.OfficeLength.Points(12),
                LeftIndent = global::OfficeIMO.OfficeLength.Points(10),
                Color = global::OfficeIMO.OfficeColor.FromRgb(55, 65, 81)
            };

            doc.Compose(compose =>
                compose.Page(page =>
                    page.Content(content =>
                        content.Row(row => {
                            row.PercentColumn(50, column => column.Bullets(new[] { "Stable bullet", "Wrapped bullet" }, style: style));
                            row.PercentColumn(50, column => column.Numbered(new[] { "First step", "Second step" }, startNumber: 3, style: style));
                        }))));

            style.FontSize = global::OfficeIMO.OfficeLength.Points(20);
            style.LeftIndent = global::OfficeIMO.OfficeLength.Points(0);

            var page = Assert.IsType<PageBlock>(Assert.Single(doc.Blocks));
            var row = Assert.IsType<RowBlock>(Assert.Single(page.Blocks));
            var bullets = Assert.IsType<BulletListBlock>(Assert.Single(row.Columns[0].Blocks));
            var numbered = Assert.IsType<NumberedListBlock>(Assert.Single(row.Columns[1].Blocks));

            Assert.Equal(new[] { "Stable bullet", "Wrapped bullet" }, bullets.Items);
            Assert.Equal(new[] { "First step", "Second step" }, numbered.Items);
            Assert.Equal(3, numbered.StartNumber);
            Assert.Equal(12, bullets.Style!.FontSize?.ToPoints());
            Assert.Equal(10, numbered.Style!.LeftIndent.ToPoints());
        }

        [Fact]
        public void RowColumn_CanComposePanelParagraph() {
            var doc = PdfDocument.Create();

            doc.Compose(compose =>
                compose.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.PercentColumn(100, column => column.PanelParagraph(
                                paragraph => paragraph.Bold("Callout").Text(": stable panel in a column"),
                                new PdfPanelStyle {
                                    Background = global::OfficeIMO.OfficeColor.FromRgb(248, 250, 252),
                                    BorderColor = global::OfficeIMO.OfficeColor.FromRgb(183, 194, 207),
                                    PaddingX = global::OfficeIMO.OfficeLength.Points(8),
                                    PaddingY = global::OfficeIMO.OfficeLength.Points(6),
                                    KeepTogether = true
                                }))))));

            var page = Assert.IsType<PageBlock>(Assert.Single(doc.Blocks));
            var row = Assert.IsType<RowBlock>(Assert.Single(page.Blocks));
            var panel = Assert.IsType<ContainerBlock>(Assert.Single(row.Columns[0].Blocks));

            Assert.True(panel.Style!.KeepTogether);
            Assert.Equal(global::OfficeIMO.OfficeColor.FromRgb(248, 250, 252), panel.Style.Background);
            Assert.Contains(Assert.IsType<RichParagraphBlock>(Assert.Single(panel.Blocks)).Runs, run => run.Bold);
        }

        [Fact]
        public void RowColumn_CanComposePanelFromCommonFlowBlocks() {
            var doc = PdfDocument.Create();

            doc.Compose(compose =>
                compose.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.PercentColumn(100, column => column.Panel(panel => panel
                                .H3("Column Panel")
                                .Paragraph(paragraph => paragraph.Text("Column-local composed panels reuse the same core primitive."))
                                .Bullets(new[] { "Stable in row flow" })))))));

            var page = Assert.IsType<PageBlock>(Assert.Single(doc.Blocks));
            var row = Assert.IsType<RowBlock>(Assert.Single(page.Blocks));
            var panel = Assert.IsType<ContainerBlock>(Assert.Single(row.Columns[0].Blocks));
            string text = PdfReadDocument.Open(doc.ToBytes()).ExtractText();

            Assert.Contains("Column Panel", text, StringComparison.Ordinal);
            Assert.Contains("Column-local composed panels", text, StringComparison.Ordinal);
            Assert.Contains("Stable in row flow", text, StringComparison.Ordinal);
        }

        [Fact]
        public void RowColumn_CanComposeTable() {
            var doc = PdfDocument.Create();
            var style = PdfTableStyles.Light();

            doc.Compose(compose =>
                compose.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.PercentColumn(100, column => column.Table(new[] {
                                new[] { "Metric", "Value" },
                                new[] { "Score", "98" }
                            }, style: style))))));

            var page = Assert.IsType<PageBlock>(Assert.Single(doc.Blocks));
            var row = Assert.IsType<RowBlock>(Assert.Single(page.Blocks));
            var table = Assert.IsType<TableBlock>(Assert.Single(row.Columns[0].Blocks));

            Assert.NotNull(table.Style);
            Assert.Equal(style.BorderColor, table.Style!.BorderColor);
            Assert.Equal(style.CellPaddingX.ToPoints(), table.Style.CellPaddingX.ToPoints());
            Assert.Equal(2, table.Rows.Count);
            Assert.Equal(new[] { "Metric", "Value" }, table.Rows[0]);
            Assert.Equal(new[] { "Score", "98" }, table.Rows[1]);
        }

        [Theory]
        [InlineData(0)]
        [InlineData(-1)]
        [InlineData(101)]
        [InlineData(double.NaN)]
        [InlineData(double.PositiveInfinity)]
        public void PercentageWidth_RejectsInvalidValue(double widthPercent) {
            Assert.Throws<ArgumentOutOfRangeException>(() => PdfColumnWidth.Percent(widthPercent));
        }

        [Fact]
        public void ColumnRejectsDefaultWidthValue() {
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfDocument.Create(document => document.Content(content => content.Row(row =>
                    row.Column(default, column => column.Text("Invalid"))))));
        }

        [Fact]
        public void RowColumn_PreservesDecoratedContent() {
            var document = PdfDocument.Create(document => document.Content(content => content.Row(row =>
                row.RelativeColumn(column => column.Element(element => element
                    .Padding(global::OfficeIMO.OfficeLength.Points(4))
                    .Content(container => container.Text("Nested")))))));
            Assert.Contains("Nested", PdfReadDocument.Open(document.ToBytes()).ExtractText());
        }
    }
}
