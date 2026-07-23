using OfficeIMO.Word.Pdf;
using OfficeIMO.Word;
using W = DocumentFormat.OpenXml.Wordprocessing;
using System.IO;
using System.Reflection;
using Xunit;

namespace OfficeIMO.Tests.Pdf {
    public class TableLayoutCacheTests {
        [Fact]
        public void ColumnWidthsAreCachedPerTable() {
            using WordDocument document = WordDocument.Create();
            WordTable table = document.AddTable(2, 2);
            table.Rows[0].Cells[0].Width = 1440;
            table.Rows[1].Cells[0].Width = 1440;

            TableLayout first = TableLayoutCache.GetLayout(table);
            TableLayout second = TableLayoutCache.GetLayout(table);

            Assert.Same(first, second);
            Assert.Equal(2, first.ColumnWidths.Length);
            Assert.Equal(72f, first.ColumnWidths[0]);
        }

        [Fact]
        public void WiderCellPreferredWidthsExpandGridColumns() {
            using WordDocument document = WordDocument.Create();
            WordTable table = document.AddTable(1, 2);
            table.GridColumnWidth = new List<int> { 720, 720 };
            table.Rows[0].Cells[0].Width = 1440;
            table.Rows[0].Cells[1].Width = 720;

            TableLayout layout = TableLayoutCache.GetLayout(table);

            Assert.Equal(72f, layout.ColumnWidths[0]);
            Assert.Equal(36f, layout.ColumnWidths[1]);
        }

        [Fact]
        public void DxaWidthOfTwentyFourHundredTwipsIsPreservedWhenAuthored() {
            using WordDocument document = WordDocument.Create();
            WordTable table = document.AddTable(1, 2);
            table.Rows[0].Cells[0].Width = 2400;
            table.Rows[0].Cells[0].WidthType = DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues.Dxa;
            table.Rows[0].Cells[1].Width = 720;
            table.Rows[0].Cells[1].WidthType = DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues.Dxa;

            TableLayout layout = TableLayoutCache.GetLayout(table);

            Assert.Equal(120f, layout.ColumnWidths[0]);
            Assert.Equal(36f, layout.ColumnWidths[1]);
        }

        [Fact]
        public void RowGridBeforeOffsetsCellsAndExpandsColumnCount() {
            using WordDocument document = WordDocument.Create();
            WordTable table = document.AddTable(2, 3);
            table.GridColumnWidth = new List<int>();
            table.Rows[1].Cells[2].Remove();
            table.Rows[1]._tableRow.TableRowProperties ??= new W.TableRowProperties();
            table.Rows[1]._tableRow.TableRowProperties.Append(new W.GridBefore { Val = 1 });
            table.Rows[1]._tableRow.TableRowProperties.Append(new W.GridAfter { Val = 1 });
            table.Rows[1].Cells[0].Width = 2880;
            table.Rows[1].Cells[0].WidthType = W.TableWidthUnitValues.Dxa;

            TableLayout layout = TableLayoutCache.GetLayout(table);

            Assert.Equal(4, layout.ColumnWidths.Length);
            Assert.Equal(0, layout.GetRowStartColumn(0));
            Assert.Equal(1, layout.GetRowStartColumn(1));
            Assert.Equal(144f, layout.ColumnWidths[1]);
        }

        [Theory]
        [InlineData(true)]
        [InlineData(false)]
        public void OversizedRowGridOffsetsAreRejectedBeforeColumnArraysAreAllocated(bool before) {
            using WordDocument document = WordDocument.Create();
            WordTable table = document.AddTable(1, 1);
            table.GridColumnWidth = new List<int>();
            table.Rows[0]._tableRow.TableRowProperties ??= new W.TableRowProperties();
            if (before) {
                table.Rows[0]._tableRow.TableRowProperties.Append(new W.GridBefore { Val = 16_385 });
            } else {
                table.Rows[0]._tableRow.TableRowProperties.Append(new W.GridAfter { Val = 16_385 });
            }

            Assert.Throws<InvalidDataException>(() => TableLayoutCache.GetLayout(table));
        }

        [Fact]
        public void CombinedRowGridOffsetsAreRejectedBeforeColumnArraysAreAllocated() {
            using WordDocument document = WordDocument.Create();
            WordTable table = document.AddTable(1, 1);
            table.GridColumnWidth = new List<int>();
            table.Rows[0]._tableRow.TableRowProperties ??= new W.TableRowProperties();
            table.Rows[0]._tableRow.TableRowProperties.Append(new W.GridBefore { Val = 16_384 });
            table.Rows[0]._tableRow.TableRowProperties.Append(new W.GridAfter { Val = 16_384 });
            table.Rows[0].Cells[0].HorizontalMerge = W.MergedCellValues.Continue;

            Assert.Throws<InvalidDataException>(() => TableLayoutCache.GetLayout(table));
        }

        [Fact]
        public void NestedTableWidthsPropagateToParent() {
            using WordDocument document = WordDocument.Create();
            WordTable outer = document.AddTable(1, 2);
            outer.Rows[0].Cells[0].Width = 1440;
            outer.Rows[0].Cells[1].Width = 720;

            WordTable inner = outer.Rows[0].Cells[1].AddTable(1, 1);
            inner.Rows[0].Cells[0].Width = 2880;

            TableLayout outerLayout = TableLayoutCache.GetLayout(outer);
            Assert.Equal(144f, outerLayout.ColumnWidths[1]);

            TableLayout innerLayout = TableLayoutCache.GetLayout(inner);
            TableLayout innerLayoutSecond = TableLayoutCache.GetLayout(inner);
            Assert.Same(innerLayout, innerLayoutSecond);
            Assert.Equal(144f, innerLayout.ColumnWidths[0]);
        }

        [Fact]
        public void RecursiveNestedTablesAreMeasured() {
            using WordDocument document = WordDocument.Create();
            WordTable outer = document.AddTable(1, 1);
            outer.Rows[0].Cells[0].Width = 720;
            WordTable middle = outer.Rows[0].Cells[0].AddTable(1, 1);
            middle.Rows[0].Cells[0].Width = 720;
            WordTable inner = middle.Rows[0].Cells[0].AddTable(1, 1);
            inner.Rows[0].Cells[0].Width = 2880;

            TableLayout outerLayout = TableLayoutCache.GetLayout(outer);
            Assert.Equal(144f, outerLayout.ColumnWidths[0]);

            TableLayout middleLayout = TableLayoutCache.GetLayout(middle);
            Assert.Equal(144f, middleLayout.ColumnWidths[0]);
            TableLayout innerLayout = TableLayoutCache.GetLayout(inner);
            Assert.Equal(144f, innerLayout.ColumnWidths[0]);
        }

        [Fact]
        public void ExcessiveNestedTableDepthIsRejectedBeforeRecursiveLayoutCanOverflow() {
            using WordDocument document = WordDocument.Create();
            WordTable root = document.AddTable(1, 1);
            WordTable current = root;
            for (int depth = 0; depth < 128; depth++) {
                current = current.Rows[0].Cells[0].AddTable(1, 1);
            }

            Assert.Throws<InvalidDataException>(() => TableLayoutCache.GetLayout(root));
        }

        [Fact]
        public void StyleTabStopInspectionIsBoundedEvenWhenEntriesAreInvalid() {
            using WordDocument document = WordDocument.Create();
            WordParagraph paragraph = document.AddParagraph("Bounded tabs");
            const string styleId = "HostileTabStops";
            var tabs = new W.Tabs();
            for (int index = 0; index < 1_024; index++) {
                tabs.Append(new W.TabStop { Position = 0, Val = W.TabStopValues.Left });
            }

            tabs.Append(new W.TabStop { Position = 1_440, Val = W.TabStopValues.Left });
            W.Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.Append(new W.Style(
                new W.StyleName { Val = styleId },
                new W.StyleParagraphProperties(tabs)) {
                Type = W.StyleValues.Paragraph,
                StyleId = styleId
            });
            paragraph._paragraph.ParagraphProperties ??= new W.ParagraphProperties();
            paragraph._paragraph.ParagraphProperties.ParagraphStyleId = new W.ParagraphStyleId { Val = styleId };

            MethodInfo method = typeof(WordPdfConverterExtensions).GetMethod(
                "GetNativeParagraphEffectiveTabStops",
                BindingFlags.NonPublic | BindingFlags.Static)!;
            var effectiveTabStops = Assert.IsAssignableFrom<IReadOnlyList<WordTabStop>>(method.Invoke(null, new object[] { paragraph }));

            Assert.Empty(effectiveTabStops);
        }
    }
}
