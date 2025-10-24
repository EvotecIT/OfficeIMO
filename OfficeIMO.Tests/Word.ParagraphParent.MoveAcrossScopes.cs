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

            // Create in header and verify parent
            var header = section.GetHeader()!;
            var headerParagraph = header.AddParagraph("Header paragraph");
            Assert.Same(header, headerParagraph.Parent);

            // Create in footer and verify parent
            var footer = section.GetFooter()!;
            var footerParagraph = footer.AddParagraph("Footer paragraph");
            Assert.Same(footer, footerParagraph.Parent);
        }

        [Fact]
        public void ParentUpdates_WhenMovingBetweenSections() {
            using var doc = WordDocument.Create();
            var s1 = doc.AddSection();
            var s2 = doc.AddSection();

            var p = s1.AddParagraph("In S1");
            Assert.Same(s1, p.Parent);

            // Reinsert existing paragraph after an anchor within section 2
            var anchor = s2.AddParagraph("Anchor in S2");
            anchor.AddParagraphAfterSelf(s2, p);
            Assert.Same(s2, p.Parent);
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

            // Move to (0,1)
            c01.AddParagraph(p);
            parentCell = Assert.IsType<WordTableCell>(p.Parent);
            Assert.True(parentCell == c01);
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
            Assert.Same(inner, parentCell.ParentTable);
            Assert.True(outerCell == outer.Rows[0].Cells[0]);
            Assert.True(outerCell.HasNestedTables);
        }
    }
}
