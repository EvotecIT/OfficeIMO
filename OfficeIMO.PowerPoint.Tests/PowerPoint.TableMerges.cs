using System;
using System.IO;
using System.Linq;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests {
    public class PowerPointTableMerges {
        private static PowerPointTable Table(PowerPointSlide slide, int rows, int columns) {
            PowerPointTable table = slide.AddTable(rows, columns);
            for (int r = 0; r < rows; r++) {
                for (int c = 0; c < columns; c++) table.GetCell(r, c).Text = r + "," + c;
            }
            return table;
        }

        /// <summary>One token per grid cell: "RxC" for an anchor, h/v/hv for covered cells and "-" for a plain cell.</summary>
        private static string Shape(PowerPointTable table) => string.Join(" / ", table.TableElement.Elements<A.TableRow>().Select(row =>
            string.Join(" ", row.Elements<A.TableCell>().Select(cell => {
                string covered = (cell.HorizontalMerge?.Value == true ? "h" : string.Empty) + (cell.VerticalMerge?.Value == true ? "v" : string.Empty);
                if (covered.Length > 0) return covered;
                int rows = cell.RowSpan?.Value ?? 1, columns = cell.GridSpan?.Value ?? 1;
                return rows > 1 || columns > 1 ? rows + "x" + columns : "-";
            }))));

        [Fact]
        public void RemovingAnAnchorRowHandsTheMergedCellToTheNextRow() {
            using var document = PowerPointDocument.Create();
            PowerPointTable table = Table(document.AddSlide(), 4, 2);
            table.MergeCells(0, 0, 2, 0);
            table.RemoveRow(0);
            Assert.Equal("2x1 - / v - / - -", Shape(table));
            Assert.Equal("0,0", table.GetCell(0, 0).Text);
        }

        [Fact]
        public void RemovingARowInsideAMergeShrinksItsSpan() {
            using var document = PowerPointDocument.Create();
            PowerPointTable table = Table(document.AddSlide(), 3, 2);
            table.MergeCells(0, 0, 1, 1);
            table.RemoveRow(1);
            Assert.Equal("1x2 h / - -", Shape(table));
        }

        [Fact]
        public void RemovingAnAnchorColumnHandsTheMergedCellToTheNextColumn() {
            using var document = PowerPointDocument.Create();
            PowerPointTable table = Table(document.AddSlide(), 2, 4);
            table.MergeCells(0, 0, 0, 2);
            table.RemoveColumn(0);
            Assert.Equal("1x2 h - / - - -", Shape(table));
            Assert.Equal("0,0", table.GetCell(0, 0).Text);
        }

        [Fact]
        public void RemovingAColumnInsideAMergeShrinksItsSpan() {
            using var document = PowerPointDocument.Create();
            PowerPointTable table = Table(document.AddSlide(), 2, 3);
            table.MergeCells(0, 0, 1, 1);
            table.RemoveColumn(1);
            Assert.Equal("2x1 - / v -", Shape(table));
        }

        [Fact]
        public void InsertedRowAndColumnJoinTheMergeTheyCross() {
            using var document = PowerPointDocument.Create();
            PowerPointTable table = Table(document.AddSlide(), 2, 2);
            table.MergeCells(0, 0, 1, 1);
            table.AddRow(1);
            table.AddColumn(1);
            Assert.Equal("3x3 h h / v hv hv / v hv hv", Shape(table));
        }

        [Fact]
        public void RowsAndColumnsCopiedFromMergedTemplatesDoNotRepeatTheirSpans() {
            using var document = PowerPointDocument.Create();
            PowerPointTable table = Table(document.AddSlide(), 2, 2);
            table.MergeCells(0, 0, 1, 0);
            table.AddRowFromTemplate(0);
            Assert.Equal("2x1 - / v - / - -", Shape(table));

            PowerPointTable wide = Table(document.Slides[0], 2, 2);
            wide.MergeCells(0, 0, 0, 1);
            wide.AddColumnFromTemplate(0);
            Assert.Equal("1x2 h - / - - -", Shape(wide));

            PowerPointTable rectangleRows = Table(document.Slides[0], 2, 2);
            rectangleRows.MergeCells(0, 0, 1, 1);
            rectangleRows.AddRowFromTemplate(0);
            Assert.Equal("2x2 h / v hv / 1x2 h", Shape(rectangleRows));

            PowerPointTable rectangleColumns = Table(document.Slides[0], 2, 2);
            rectangleColumns.MergeCells(0, 0, 1, 1);
            rectangleColumns.AddColumnFromTemplate(0);
            Assert.Equal("2x2 h 2x1 / v hv v", Shape(rectangleColumns));
        }

        [Fact]
        public void MergingOverAnExistingMergeReleasesWhatFallsOutsideIt() {
            using var document = PowerPointDocument.Create();
            PowerPointTable table = Table(document.AddSlide(), 3, 3);
            table.MergeCells(0, 0, 1, 1);
            table.MergeCells(1, 1, 2, 2);
            Assert.Equal("- - - / - 2x2 h / - v hv", Shape(table));
        }
    }
}
