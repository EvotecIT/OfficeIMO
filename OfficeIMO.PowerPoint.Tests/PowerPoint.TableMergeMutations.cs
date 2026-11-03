using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests {
    public class PowerPointTableMergeMutations {
        // Cells are written as "RxC" for an anchor spanning R rows and C columns,
        // "h"/"v"/"hv" for covered cells and "." for ordinary cells.
        [Theory]
        [InlineData(2, 3, 0, 0, 0, 1, "RemoveColumn", 1, "Merged . / . .")]
        [InlineData(2, 3, 0, 0, 0, 1, "RemoveColumn", 0, "Merged . / . .")]
        [InlineData(2, 3, 0, 0, 0, 1, "AddColumn", 1, "1x3 h h . / . . . .")]
        [InlineData(2, 3, 0, 0, 0, 1, "AddColumn", 2, "1x2 h . . / . . . .")]
        [InlineData(2, 3, 0, 0, 0, 1, "AddColumn", 3, "1x2 h . . / . . . .")]
        [InlineData(2, 3, 0, 0, 0, 1, "AddColumn", 0, ". 1x2 h . / . . . .")]
        [InlineData(2, 3, 0, 0, 0, 1, "AddColumnFromTemplate", 1, "1x3 h h . / . . . .")]
        [InlineData(3, 2, 0, 0, 1, 0, "RemoveRow", 1, "Merged . / . .")]
        [InlineData(3, 2, 0, 0, 1, 0, "RemoveRow", 0, "Merged . / . .")]
        [InlineData(3, 2, 0, 0, 1, 0, "RemoveRowWrapper", 0, "Merged . / . .")]
        [InlineData(3, 2, 0, 0, 1, 0, "AddRow", 1, "3x1 . / v . / v . / . .")]
        [InlineData(3, 2, 0, 0, 1, 0, "AddRow", 2, "2x1 . / v . / . . / . .")]
        [InlineData(3, 2, 0, 0, 1, 0, "AddRow", 0, ". . / 2x1 . / v . / . .")]
        [InlineData(3, 2, 0, 0, 1, 0, "AddRowFromTemplate", 1, "3x1 . / v . / v . / . .")]
        [InlineData(3, 3, 0, 0, 1, 1, "RemoveColumn", 0, "2x1 . / v . / . .")]
        [InlineData(3, 3, 0, 0, 1, 1, "RemoveRow", 0, "1x2 h . / . . .")]
        [InlineData(3, 3, 1, 1, 2, 2, "RemoveColumn", 0, ". . / 2x2 h / v hv")]
        public void RowAndColumnMutationsKeepMergedRangesConsistent(int rows, int columns,
            int startRow, int startColumn, int endRow, int endColumn,
            string operation, int index, string expected) {
            using var stream = new MemoryStream();
            using (PowerPointDocument presentation = PowerPointDocument.Create(stream)) {
                PowerPointTable table = presentation.AddSlide().AddTable(rows, columns);
                PowerPointTableCell anchorHandle = table.GetCell(startRow, startColumn);
                anchorHandle.Text = "Merged";
                table.MergeCells(startRow, startColumn, endRow, endColumn);

                switch (operation) {
                    case "RemoveColumn": table.RemoveColumn(index); break;
                    case "AddColumn": table.AddColumn(index); break;
                    case "AddColumnFromTemplate": table.AddColumnFromTemplate(columns - 1, index); break;
                    case "RemoveRow": table.RemoveRow(index); break;
                    case "RemoveRowWrapper": table.GetRow(index).Remove(); break;
                    case "AddRow": table.AddRow(index); break;
                    case "AddRowFromTemplate": table.AddRowFromTemplate(rows - 1, index); break;
                    default: throw new ArgumentOutOfRangeException(nameof(operation));
                }

                anchorHandle.Text = "Updated";
                Assert.Empty(presentation.ValidateDocument());
                presentation.Save();
            }

            stream.Position = 0;
            using PresentationDocument document = PresentationDocument.Open(stream, false);
            A.Table saved = document.PresentationPart!.SlideParts.Single().Slide.Descendants<A.Table>().Single();
            int savedColumns = saved.TableGrid!.Elements<A.GridColumn>().Count();
            Assert.All(saved.Elements<A.TableRow>(), _ => Assert.Equal(savedColumns, _.Elements<A.TableCell>().Count()));

            // A range that collapses to one cell reports the surviving anchor text instead of a span.
            string actual = string.Join(" / ", saved.Elements<A.TableRow>()
                .Select(_ => string.Join(" ", _.Elements<A.TableCell>().Select(Describe))));
            if (expected.StartsWith("Merged", StringComparison.Ordinal)) {
                A.TableCell anchor = saved.Elements<A.TableRow>().ElementAt(startRow).Elements<A.TableCell>().ElementAt(startColumn);
                Assert.Equal("Updated", anchor.InnerText);
                expected = "." + expected.Substring("Merged".Length);
            } else {
                A.TableCell anchor = saved.Descendants<A.TableCell>().Single(_ => _.GridSpan != null || _.RowSpan != null);
                Assert.Equal("Updated", anchor.InnerText);
            }
            Assert.Equal(expected, actual);
        }

        [Theory]
        [InlineData(true, 0, false, "1x2 h . / . . . / 1x2 h . / . . .")]
        [InlineData(true, 3, true, ". . . / 1x2 h . / . . . / 1x2 h .")]
        [InlineData(false, 0, false, "2x1 . 2x1 . / v . v . / . . . .")]
        [InlineData(false, 3, true, ". 2x1 . 2x1 / . v . v / . . . .")]
        public void TemplateMutationsPreserveClonedMerges(bool cloneRow, int index, bool clearText, string expected) {
            using var stream = new MemoryStream();
            using (PowerPointDocument presentation = PowerPointDocument.Create(stream)) {
                PowerPointTable table = presentation.AddSlide().AddTable(3, 3);
                table.GetCell(cloneRow ? 1 : 0, cloneRow ? 0 : 1).Text = "Template";
                if (cloneRow) {
                    table.MergeCells(1, 0, 1, 1);
                    table.AddRowFromTemplate(1, index, clearText);
                } else {
                    table.MergeCells(0, 1, 1, 1);
                    table.AddColumnFromTemplate(1, index, clearText);
                }

                Assert.Empty(presentation.ValidateDocument());
                presentation.Save();
            }

            stream.Position = 0;
            using PresentationDocument document = PresentationDocument.Open(stream, false);
            A.Table saved = document.PresentationPart!.SlideParts.Single().Slide.Descendants<A.Table>().Single();
            Assert.Equal(expected, string.Join(" / ", saved.Elements<A.TableRow>()
                .Select(row => string.Join(" ", row.Elements<A.TableCell>().Select(Describe)))));
            int originalIndex = index <= 1 ? 2 : 1;
            A.TableCell original = saved.Elements<A.TableRow>().ElementAt(cloneRow ? originalIndex : 0)
                .Elements<A.TableCell>().ElementAt(cloneRow ? 0 : originalIndex);
            A.TableCell clone = saved.Elements<A.TableRow>().ElementAt(cloneRow ? index : 0)
                .Elements<A.TableCell>().ElementAt(cloneRow ? 0 : index);
            Assert.Equal("Template", original.InnerText);
            Assert.Equal(clearText ? string.Empty : "Template", clone.InnerText);
        }

        [Theory]
        [InlineData(true, "3x1 . . / v . . / v . . / . . . / 1x2 h .")]
        [InlineData(false, "1x3 h h . 2x1 / . . . . v / . . . . .")]
        public void TemplateMergesDoNotOverlapRangesExpandedByInsertion(bool cloneRow, string expected) {
            using var stream = new MemoryStream();
            using (PowerPointDocument presentation = PowerPointDocument.Create(stream)) {
                PowerPointTable table = presentation.AddSlide().AddTable(cloneRow ? 4 : 3, cloneRow ? 3 : 4);
                if (cloneRow) {
                    table.MergeCells(0, 0, 1, 0);
                    table.MergeCells(3, 0, 3, 1);
                    table.AddRowFromTemplate(3, 1);
                } else {
                    table.MergeCells(0, 0, 0, 1);
                    table.MergeCells(0, 3, 1, 3);
                    table.AddColumnFromTemplate(3, 1);
                }

                Assert.Empty(presentation.ValidateDocument());
                presentation.Save();
            }

            stream.Position = 0;
            using PresentationDocument document = PresentationDocument.Open(stream, false);
            A.Table saved = document.PresentationPart!.SlideParts.Single().Slide.Descendants<A.Table>().Single();
            Assert.Equal(expected, string.Join(" / ", saved.Elements<A.TableRow>()
                .Select(row => string.Join(" ", row.Elements<A.TableCell>().Select(Describe)))));
        }

        private static string Describe(A.TableCell cell) {
            if (cell.GridSpan != null || cell.RowSpan != null) {
                return $"{cell.RowSpan?.Value ?? 1}x{cell.GridSpan?.Value ?? 1}";
            }
            string merge = (cell.HorizontalMerge?.Value == true ? "h" : string.Empty)
                + (cell.VerticalMerge?.Value == true ? "v" : string.Empty);
            return merge.Length == 0 ? "." : merge;
        }
    }
}
