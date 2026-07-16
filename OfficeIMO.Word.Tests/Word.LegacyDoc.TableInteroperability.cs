using OfficeIMO.Drawing.Internal;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void LegacyDoc_SaveDocPath_PreservesOuterTableCellPapxBoundariesWhenCommentsArePresent() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");

            try {
                using (WordDocument document = WordDocument.Create()) {
                    document.AddParagraph("Commented").AddComment("OfficeIMO", "OI", "Comment");
                    WordTable table = document.AddTable(1, 2);
                    table.Rows[0].Cells[0].AddParagraph("A1", removeExistingParagraphs: true);
                    table.Rows[0].Cells[1].AddParagraph("B1", removeExistingParagraphs: true);
                    document.Save(docPath);
                }

                Assert.True(
                    OfficeCompoundFileReader.TryRead(File.ReadAllBytes(docPath), out OfficeCompoundFile? compoundFile, out string? error),
                    error);
                byte[] wordDocumentStream = compoundFile!.Streams["WordDocument"];

                Assert.Equal(
                    3,
                    CountBytePattern(wordDocumentStream, 0x49, 0x66, 0x01, 0x00, 0x00, 0x00));

                using WordDocument reloaded = WordDocument.Load(docPath);
                WordTable reloadedTable = Assert.Single(reloaded.Tables);
                WordTableRow row = Assert.Single(reloadedTable.Rows);
                Assert.Equal("A1", row.Cells[0].Paragraphs[0].Text);
                Assert.Equal("B1", row.Cells[1].Paragraphs[0].Text);
            } finally {
                if (File.Exists(docPath)) {
                    File.Delete(docPath);
                }
            }
        }

        private static int CountBytePattern(byte[] bytes, params byte[] pattern) {
            int count = 0;
            for (int offset = 0; offset <= bytes.Length - pattern.Length; offset++) {
                if (bytes.AsSpan(offset, pattern.Length).SequenceEqual(pattern)) {
                    count++;
                }
            }

            return count;
        }
    }
}
