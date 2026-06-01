using System;
using System.Linq;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf {
    public class PdfRowComposeTests {
        [Fact]
        public void RowRejectsEmptyComposition() {
            var doc = PdfDoc.Create();

            var exception = Assert.Throws<InvalidOperationException>(() =>
                doc.Compose(compose =>
                    compose.Page(page =>
                        page.Content(content =>
                            content.Row(_ => { })))));

            Assert.Contains("Rows require at least one column.", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void ColumnRejectsNonPositiveWidth() {
            var doc = PdfDoc.Create();

            Assert.Throws<ArgumentOutOfRangeException>(() =>
                doc.Compose(compose =>
                    compose.Page(page =>
                        page.Content(content =>
                            content.Row(row =>
                                row.Column(0, _ => { }))))));
        }

        [Theory]
        [InlineData(double.NaN)]
        [InlineData(double.PositiveInfinity)]
        [InlineData(double.NegativeInfinity)]
        public void ColumnRejectsNonFiniteWidth(double widthPercent) {
            var doc = PdfDoc.Create();

            var exception = Assert.Throws<ArgumentOutOfRangeException>(() =>
                doc.Compose(compose =>
                    compose.Page(page =>
                        page.Content(content =>
                            content.Row(row =>
                                row.Column(widthPercent, _ => { }))))));

            Assert.Equal("widthPercent", exception.ParamName);
            Assert.Contains("Column width must be a finite percentage.", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void ColumnRejectsWidthOverOneHundred() {
            var doc = PdfDoc.Create();

            Assert.Throws<ArgumentOutOfRangeException>(() =>
                doc.Compose(compose =>
                    compose.Page(page =>
                        page.Content(content =>
                            content.Row(row =>
                                row.Column(150, _ => { }))))));
        }

        [Fact]
        public void ColumnRejectsWhenTotalWouldExceedOneHundred() {
            var doc = PdfDoc.Create();

            Assert.Throws<InvalidOperationException>(() =>
                doc.Compose(compose =>
                    compose.Page(page =>
                        page.Content(content =>
                            content.Row(row => {
                                row.Column(60, _ => { });
                                row.Column(50, _ => { });
                            })))));
        }

        [Fact]
        public void ColumnsAreNormalizedWhenTotalLessThanOneHundred() {
            var doc = PdfDoc.Create();

            doc.Compose(compose =>
                compose.Page(page =>
                    page.Content(content =>
                        content.Row(row => {
                            row.Column(30, _ => { });
                            row.Column(20, _ => { });
                        }))));

            var page = Assert.IsType<PageBlock>(Assert.Single(doc.Blocks));
            var row = Assert.IsType<RowBlock>(Assert.Single(page.Blocks));
            Assert.Equal(2, row.Columns.Count);

            var widths = row.Columns.Select(c => c.WidthPercent).ToArray();
            Assert.InRange(widths[0], 59.9, 60.1);
            Assert.InRange(widths[1], 39.9, 40.1);
            var total = widths.Sum();
            Assert.InRange(total, 99.99, 100.01);
        }

        [Fact]
        public void RowGap_IsStoredOnModel() {
            var doc = PdfDoc.Create();

            doc.Compose(compose =>
                compose.Page(page =>
                    page.Content(content =>
                        content.Row(row => {
                            row.Gap(18);
                            row.Column(50, _ => { });
                            row.Column(50, _ => { });
                        }))));

            var page = Assert.IsType<PageBlock>(Assert.Single(doc.Blocks));
            var row = Assert.IsType<RowBlock>(Assert.Single(page.Blocks));
            Assert.Equal(18, row.Gap);
        }

        [Fact]
        public void RowStyle_IsStoredOnModelAndSnapshotsInput() {
            var doc = PdfDoc.Create();
            var style = new PdfRowStyle {
                Gap = 18,
                SpacingBefore = 7,
                SpacingAfter = 9,
                ColumnSeparatorColor = new PdfColor(0.12, 0.34, 0.56),
                ColumnSeparatorWidth = 1.25,
                KeepTogether = true,
                KeepWithNext = true
            };

            doc.Compose(compose =>
                compose.Page(page =>
                    page.Content(content =>
                        content.Row(row => {
                            row.Style(style);
                            row.Column(50, column => column.Paragraph(paragraph => paragraph.Text("Left")));
                            row.Column(50, column => column.Paragraph(paragraph => paragraph.Text("Right")));
                        }))));

            style.Gap = 4;
            style.SpacingBefore = 1;
            style.SpacingAfter = 2;
            style.ColumnSeparatorColor = PdfColor.Black;
            style.ColumnSeparatorWidth = 0.5;
            style.KeepTogether = false;
            style.KeepWithNext = false;

            var page = Assert.IsType<PageBlock>(Assert.Single(doc.Blocks));
            var row = Assert.IsType<RowBlock>(Assert.Single(page.Blocks));

            Assert.Equal(18, row.Gap);
            Assert.Equal(18, row.Style!.Gap);
            Assert.Equal(7, row.Style.SpacingBefore);
            Assert.Equal(9, row.Style.SpacingAfter);
            Assert.Equal(new PdfColor(0.12, 0.34, 0.56), row.Style.ColumnSeparatorColor);
            Assert.Equal(1.25, row.Style.ColumnSeparatorWidth);
            Assert.True(row.Style.KeepTogether);
            Assert.True(row.Style.KeepWithNext);
        }

        [Fact]
        public void RowUsesBuiltInWordLikeGapUnlessExplicitlyOptedOut() {
            var defaultDoc = PdfDoc.Create();
            defaultDoc.Compose(compose =>
                compose.Page(page =>
                    page.Content(content =>
                        content.Row(row => {
                            row.Column(50, column => column.Paragraph(paragraph => paragraph.Text("Left")));
                            row.Column(50, column => column.Paragraph(paragraph => paragraph.Text("Right")));
                        }))));

            var defaultPage = Assert.IsType<PageBlock>(Assert.Single(defaultDoc.Blocks));
            var defaultRow = Assert.IsType<RowBlock>(Assert.Single(defaultPage.Blocks));

            Assert.Equal(PdfRowStyle.DefaultGap, defaultRow.Gap);

            var optOutDoc = PdfDoc.Create();
            optOutDoc.Compose(compose =>
                compose.Page(page =>
                    page.Content(content =>
                        content.Row(row => {
                            row.Gap(0);
                            row.Column(50, column => column.Paragraph(paragraph => paragraph.Text("Left")));
                            row.Column(50, column => column.Paragraph(paragraph => paragraph.Text("Right")));
                        }))));

            var optOutPage = Assert.IsType<PageBlock>(Assert.Single(optOutDoc.Blocks));
            var optOutRow = Assert.IsType<RowBlock>(Assert.Single(optOutPage.Blocks));

            Assert.Equal(0, optOutRow.Gap);
        }

        [Theory]
        [InlineData(-1)]
        [InlineData(double.NaN)]
        [InlineData(double.PositiveInfinity)]
        [InlineData(double.NegativeInfinity)]
        public void RowGap_RejectsInvalidValues(double gap) {
            var doc = PdfDoc.Create();

            var exception = Assert.Throws<ArgumentOutOfRangeException>(() =>
                doc.Compose(compose =>
                    compose.Page(page =>
                        page.Content(content =>
                            content.Row(row => row.Gap(gap))))));

            Assert.Equal("gap", exception.ParamName);
        }

        [Fact]
        public void RowStyle_RejectsInvalidValues() {
            Assert.Throws<ArgumentException>(() => new PdfRowStyle { Gap = double.NaN });
            Assert.Throws<ArgumentException>(() => new PdfRowStyle { Gap = -1 });
            Assert.Throws<ArgumentException>(() => new PdfRowStyle { SpacingBefore = double.PositiveInfinity });
            Assert.Throws<ArgumentException>(() => new PdfRowStyle { SpacingAfter = -1 });
            Assert.Throws<ArgumentException>(() => new PdfRowStyle { ColumnSeparatorWidth = double.NaN });
            Assert.Throws<ArgumentException>(() => new PdfRowStyle { ColumnSeparatorWidth = -1 });

            Assert.Throws<ArgumentNullException>(() =>
                PdfDoc.Create().Compose(compose =>
                    compose.Page(page =>
                        page.Content(content =>
                            content.Row(row => row.Style(null!))))));
        }

        [Fact]
        public void RowColumnSeparator_RendersBetweenColumns() {
            byte[] bytes = PdfDoc.Create(new PdfOptions {
                    PageWidth = 360,
                    PageHeight = 180,
                    MarginLeft = 30,
                    MarginRight = 30,
                    MarginTop = 30,
                    MarginBottom = 30,
                    DefaultFontSize = 10
                })
                .Compose(compose =>
                    compose.Page(page =>
                        page.Content(content =>
                            content.Row(row => row
                                .Gap(20)
                                .ColumnSeparator(new PdfColor(0.12, 0.34, 0.56), 1.25)
                                .Column(50, column => column.Paragraph(paragraph => paragraph.Text("LeftSeparatorMarker")))
                                .Column(50, column => column.Paragraph(paragraph => paragraph.Text("RightSeparatorMarker")))))))
                .ToBytes();

            string rawPdf = Encoding.ASCII.GetString(bytes);

            Assert.Contains("0.12 0.34 0.56 RG", rawPdf, StringComparison.Ordinal);
            Assert.Contains("1.25 w", rawPdf, StringComparison.Ordinal);
            Assert.Contains("180 150 m 180 ", rawPdf, StringComparison.Ordinal);
        }

        [Fact]
        public void RowGap_RejectsWhenGapsExceedContentWidthDuringRender() {
            var doc = PdfDoc.Create(new PdfOptions {
                PageWidth = 120,
                PageHeight = 160,
                MarginLeft = 20,
                MarginRight = 20,
                MarginTop = 20,
                MarginBottom = 20
            }).Compose(compose =>
                compose.Page(page =>
                    page.Content(content =>
                        content.Row(row => {
                            row.Gap(90);
                            row.Column(50, column => column.Paragraph(paragraph => paragraph.Text("Left")));
                            row.Column(50, column => column.Paragraph(paragraph => paragraph.Text("Right")));
                        }))));

            var exception = Assert.Throws<ArgumentException>(() => doc.ToBytes());
            Assert.Contains("Row column gaps must be smaller than the available page content width.", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void RowModel_ExposesReadOnlyColumnAndBlockCollections() {
            var doc = PdfDoc.Create();

            doc.Compose(compose =>
                compose.Page(page =>
                    page.Content(content =>
                        content.Row(row => {
                            row.Column(40, column => column
                                .H1("Left")
                                .Paragraph(paragraph => paragraph.Text("Body")));
                            row.Column(60, column => column.H2("Right"));
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
        public void RowColumn_ItemProvidesWordLikeFlowGroups() {
            var doc = PdfDoc.Create();

            doc.Compose(compose =>
                compose.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column => column
                                .Item(item => item
                                    .H2("Grouped heading")
                                    .Paragraph(paragraph => paragraph.Text("Grouped paragraph")))
                                .Spacer(6)
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
                PdfDoc.Create().Compose(compose =>
                    compose.Page(page =>
                        page.Content(content =>
                            content.Row(row => row.Column(100, column => column.Item(null!)))))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDoc.Create().Compose(compose =>
                    compose.Page(page =>
                        page.Content(content =>
                            content.Row(row => row.Column(100, column => column.Paragraph(null!)))))));
        }

        [Fact]
        public void RowColumn_CanComposeBulletAndNumberedLists() {
            var doc = PdfDoc.Create();
            var style = new PdfListStyle {
                FontSize = 12,
                LeftIndent = 10,
                Color = PdfColor.FromRgb(55, 65, 81)
            };

            doc.Compose(compose =>
                compose.Page(page =>
                    page.Content(content =>
                        content.Row(row => {
                            row.Column(50, column => column.Bullets(new[] { "Stable bullet", "Wrapped bullet" }, style: style));
                            row.Column(50, column => column.Numbered(new[] { "First step", "Second step" }, startNumber: 3, style: style));
                        }))));

            style.FontSize = 20;
            style.LeftIndent = 0;

            var page = Assert.IsType<PageBlock>(Assert.Single(doc.Blocks));
            var row = Assert.IsType<RowBlock>(Assert.Single(page.Blocks));
            var bullets = Assert.IsType<BulletListBlock>(Assert.Single(row.Columns[0].Blocks));
            var numbered = Assert.IsType<NumberedListBlock>(Assert.Single(row.Columns[1].Blocks));

            Assert.Equal(new[] { "Stable bullet", "Wrapped bullet" }, bullets.Items);
            Assert.Equal(new[] { "First step", "Second step" }, numbered.Items);
            Assert.Equal(3, numbered.StartNumber);
            Assert.Equal(12, bullets.Style!.FontSize);
            Assert.Equal(10, numbered.Style!.LeftIndent);
        }

        [Fact]
        public void RowColumn_CanComposePanelParagraph() {
            var doc = PdfDoc.Create();

            doc.Compose(compose =>
                compose.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column => column.PanelParagraph(
                                paragraph => paragraph.Bold("Callout").Text(": stable panel in a column"),
                                new PanelStyle {
                                    Background = PdfColor.FromRgb(248, 250, 252),
                                    BorderColor = PdfColor.FromRgb(183, 194, 207),
                                    PaddingX = 8,
                                    PaddingY = 6,
                                    KeepTogether = true
                                }))))));

            var page = Assert.IsType<PageBlock>(Assert.Single(doc.Blocks));
            var row = Assert.IsType<RowBlock>(Assert.Single(page.Blocks));
            var panel = Assert.IsType<PanelParagraphBlock>(Assert.Single(row.Columns[0].Blocks));

            Assert.True(panel.Style!.KeepTogether);
            Assert.Equal(PdfColor.FromRgb(248, 250, 252), panel.Style.Background);
            Assert.Contains(panel.Runs, run => run.Bold);
        }

        [Fact]
        public void RowColumn_CanComposeTable() {
            var doc = PdfDoc.Create();
            var style = TableStyles.Light();

            doc.Compose(compose =>
                compose.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column => column.Table(new[] {
                                new[] { "Metric", "Value" },
                                new[] { "Score", "98" }
                            }, style: style))))));

            var page = Assert.IsType<PageBlock>(Assert.Single(doc.Blocks));
            var row = Assert.IsType<RowBlock>(Assert.Single(page.Blocks));
            var table = Assert.IsType<TableBlock>(Assert.Single(row.Columns[0].Blocks));

            Assert.NotNull(table.Style);
            Assert.Equal(style.BorderColor, table.Style!.BorderColor);
            Assert.Equal(style.CellPaddingX, table.Style.CellPaddingX);
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
        public void RowColumn_RejectsInvalidWidthAtModelConstruction(double widthPercent) {
            Assert.Throws<ArgumentOutOfRangeException>(() => new RowColumn(widthPercent));
        }
    }
}
