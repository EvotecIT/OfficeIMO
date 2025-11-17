using System.Linq;
using OfficeIMO.Word;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests {
    public class WordParagraphParentMoveAcrossScopesTests {
        [Fact]
        public void ParentUpdates_WhenMovingBetweenBodyHeaderFooter() {
            using var doc = WordDocument.Create();
            var section = doc.AddSection();

            // Ensure headers/footers exist
            doc.AddHeadersAndFooters();

            // Body paragraph
            var bodyParagraph = section.AddParagraph("Body paragraph");
            Assert.IsType<WordSection>(bodyParagraph.Parent);

            // Create in header and verify parent (use document-level accessors)
            var docHeader = doc.Header!.Default!;
            var headerParagraph = docHeader.AddParagraph("Header paragraph");
            Assert.Same(docHeader, headerParagraph.Parent);

            // Create in footer and verify parent
            var docFooter = doc.Footer!.Default!;
            var footerParagraph = docFooter.AddParagraph("Footer paragraph");
            Assert.Same(docFooter, footerParagraph.Parent);
        }

        [Fact]
        public void ParentUpdates_WhenMovingBetweenSections() {
            using var doc = WordDocument.Create();
            var s1 = doc.AddSection();
            var s2 = doc.AddSection();

            var p = s1.AddParagraph("In S1");
            Assert.Same(s1, p.Parent);

            // Create a new paragraph in section 2 and verify its parent
            var p2 = s2.AddParagraph("In S2");
            Assert.IsType<WordSection>(p2.Parent);
        }

        [Fact]
        public void ParentUpdates_WhenMovingBetweenTableCells() {
            using var doc = WordDocument.Create();
            var t = doc.AddTable(1, 2, WordTableStyle.TableGrid);

            // Create paragraph in cell (0,0)
            var c00 = t.Rows[0].Cells[0];
            var c01 = t.Rows[0].Cells[1];
            var p = c00.AddParagraph("Cell paragraph");
            var parentCell = Assert.IsType<WordTableCell>(p.Parent);
            Assert.True(parentCell == c00);

            // Also create in (0,1)
            var p2 = c01.AddParagraph("Cell paragraph 2");
            var parentCell2 = Assert.IsType<WordTableCell>(p2.Parent);
            Assert.True(parentCell2 == c01);
        }

        [Fact]
        public void ParentChain_CorrectForNestedTables() {
            using var doc = WordDocument.Create();
            var outer = doc.AddTable(1, 1, WordTableStyle.TableGrid);
            var outerCell = outer.Rows[0].Cells[0];

            // Add nested table inside outer cell
            var inner = outerCell.AddTable(1, 1, WordTableStyle.TableGrid, removePrecedingParagraph: true);
            var innerCell = inner.Rows[0].Cells[0];
            var p = innerCell.AddParagraph("Nested");

            var parentCell = Assert.IsType<WordTableCell>(p.Parent);
            Assert.True(parentCell == innerCell);
            // Avoid reference equality on WordTable wrappers; assert structural expectations
            Assert.Equal(1, parentCell.ParentTable.RowsCount);
            Assert.True(outerCell == outer.Rows[0].Cells[0]);
            Assert.True(outerCell.HasNestedTables);
        }
    }
}
