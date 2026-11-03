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
                table.GetCell(startRow, startColumn).Text = "Merged";
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
                Assert.Equal("Merged", anchor.InnerText);
                expected = "." + expected.Substring("Merged".Length);
            } else {
                A.TableCell anchor = saved.Descendants<A.TableCell>().Single(_ => _.GridSpan != null || _.RowSpan != null);
                Assert.Equal("Merged", anchor.InnerText);
            }
            Assert.Equal(expected, actual);
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
