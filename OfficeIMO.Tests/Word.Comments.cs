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
                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "Test_AddReadRemoveComment.docx"))) {
                Assert.True(document.Comments.Count == 1);
                var comment = document.Comments.First();
                comment.Remove();
                document.Save(false);
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
                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "Test_RemoveAllComments.docx"))) {
                Assert.Equal(2, document.Comments.Count);
                document.RemoveAllComments();
                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "Test_RemoveAllComments.docx"))) {
                Assert.Empty(document.Comments);
            }
        }

        [Fact]
        public void Test_TrackCommentsViaDocumentProperty() {
            string filePath = Path.Combine(_directoryWithFiles, "Test_TrackComments2.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.TrackComments = true;
                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "Test_TrackComments2.docx"))) {
                Assert.True(document.TrackComments);
                document.TrackComments = false;
                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "Test_TrackComments2.docx"))) {
                Assert.False(document.TrackComments);
            }
        }

        [Fact]
        public void Test_TrackCommentsSetting() {
            string filePath = Path.Combine(_directoryWithFiles, "Test_TrackComments.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.Settings.TrackComments = true;
                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "Test_TrackComments.docx"))) {
                Assert.True(document.Settings.TrackComments);
                document.Settings.TrackComments = false;
                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "Test_TrackComments.docx"))) {
                Assert.False(document.Settings.TrackComments);
            }
        }
    }
}
