using System;
using System.IO;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void ParagraphInsertionRejectsForeignDocumentParagraphs() {
            string filePath1 = Path.Combine(_directoryWithFiles, "ParagraphGuard_Document1.docx");
            string filePath2 = Path.Combine(_directoryWithFiles, "ParagraphGuard_Document2.docx");

            using var document1 = WordDocument.Create(filePath1);
            using var document2 = WordDocument.Create(filePath2);

            var hostParagraph = document1.AddParagraph("Host");
            var foreignParagraph = document2.AddParagraph("Foreign");

            var paragraphException = Assert.Throws<InvalidOperationException>(() => hostParagraph.AddParagraph(foreignParagraph));
            Assert.Contains("different document", paragraphException.Message, StringComparison.OrdinalIgnoreCase);

            var documentException = Assert.Throws<InvalidOperationException>(() => document1.AddParagraph(foreignParagraph));
            Assert.Contains("different document", documentException.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void TableCellRejectsParagraphsFromOtherDocuments() {
            string hostPath = Path.Combine(_directoryWithFiles, "ParagraphGuard_TableHost.docx");
            string foreignPath = Path.Combine(_directoryWithFiles, "ParagraphGuard_TableForeign.docx");

            using var hostDocument = WordDocument.Create(hostPath);
            using var foreignDocument = WordDocument.Create(foreignPath);

            var table = hostDocument.AddTable(1, 1);
            var hostCell = table.Rows[0].Cells[0];
            var foreignParagraph = foreignDocument.AddParagraph("Foreign paragraph");

            var exception = Assert.Throws<InvalidOperationException>(() => hostCell.AddParagraph(foreignParagraph));
            Assert.Contains("different document", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void ParagraphInsertionRequiresMatchingSections() {
            string filePath = Path.Combine(_directoryWithFiles, "ParagraphGuard_Sections.docx");

            using var document = WordDocument.Create(filePath);
            var firstSectionParagraph = document.AddParagraph("Section 1 paragraph");
            var secondSection = document.AddSection();
            var secondSectionParagraph = secondSection.AddParagraph("Section 2 paragraph");

            var exception = Assert.Throws<InvalidOperationException>(() => secondSectionParagraph.AddParagraph(firstSectionParagraph));
            Assert.Contains("different section", exception.Message, StringComparison.OrdinalIgnoreCase);
        }
    }
}
