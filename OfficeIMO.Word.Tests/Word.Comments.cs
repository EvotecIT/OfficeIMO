using System.IO;
using System.Linq;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_AddReadRemoveComment() {
            string filePath = Path.Combine(_directoryWithFiles, "Test_AddReadRemoveComment.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Paragraph");
                document.Paragraphs[0].AddComment("John Doe", "JD", "Sample comment");
                document.Save();
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "Test_AddReadRemoveComment.docx"))) {
                Assert.True(document.Comments.Count == 1);
                var comment = document.Comments.First();
                comment.Remove();
                document.Save();
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "Test_AddReadRemoveComment.docx"))) {
                Assert.True(document.Comments.Count == 0);
            }
        }

        [Fact]
        public void Test_RemoveAllCommentsMethod() {
            string filePath = Path.Combine(_directoryWithFiles, "Test_RemoveAllComments.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Paragraph");
                document.Paragraphs[0].AddComment("John Doe", "JD", "First");
                document.Paragraphs[0].AddComment("John Doe", "JD", "Second");
                document.Save();
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "Test_RemoveAllComments.docx"))) {
                Assert.Equal(2, document.Comments.Count);
                document.RemoveAllComments();
                document.Save();
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "Test_RemoveAllComments.docx"))) {
                Assert.Empty(document.Comments);
            }
        }

    }
}
