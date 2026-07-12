using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public async Task SaveAsync_PathlessStreamBackedDocumentCopiesEditsToOriginalStream() {
            using var source = new MemoryStream();
            using var document = WordDocument.Create(source);
            document.AddParagraph("Explicit async save");

            await document.SaveAsync();

            Assert.True(source.CanRead);
            source.Position = 0;
            using WordDocument reopened = WordDocument.Load(source, readOnly: true);
            Assert.Equal("Explicit async save", Assert.Single(reopened.Paragraphs).Text);
        }

        [Fact]
        public async Task LoadAsync_StreamAutoSaveCopiesEditsBackOnAsyncDispose() {
            using var source = new MemoryStream();
            using (WordDocument created = WordDocument.Create(source)) {
                created.AddParagraph("Before");
                created.Save();
            }

            source.Position = 0;
            await using (WordDocument loaded = await WordDocument.LoadAsync(source, autoSave: true)) {
                Assert.Single(loaded.Paragraphs).SetText("After");
            }

            Assert.True(source.CanRead);
            source.Position = 0;
            using WordDocument reopened = WordDocument.Load(source, readOnly: true);
            Assert.Equal("After", Assert.Single(reopened.Paragraphs).Text);
        }

        [Fact]
        public async Task LoadAsync_NonSeekableStreamDoesNotBecomePathlessSaveTarget() {
            byte[] sourceBytes;
            using (WordDocument created = WordDocument.Create()) {
                created.AddParagraph("Buffered source");
                sourceBytes = created.ToDocx();
            }

            using var source = new NonSeekableReadStream(sourceBytes);
            using WordDocument loaded = await WordDocument.LoadAsync(source);
            loaded.AddParagraph("Unsaved edit");

            InvalidOperationException syncException = Assert.Throws<InvalidOperationException>(() => loaded.Save());
            Assert.Contains("not associated with a file path", syncException.Message, StringComparison.OrdinalIgnoreCase);

            InvalidOperationException asyncException = await Assert.ThrowsAsync<InvalidOperationException>(() => loaded.SaveAsync());
            Assert.Contains("not associated with a file path", asyncException.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void SaveCopy_StreamDoesNotRedirectLaterSourceSavesToTheCopy() {
            using var source = new MemoryStream();
            using var document = WordDocument.Create(source);
            document.AddParagraph("Shared");
            document.Save();

            using var copyStream = new MemoryStream();
            using (WordDocument copy = document.SaveCopy(copyStream)) {
                Assert.Equal("Shared", Assert.Single(copy.Paragraphs).Text);
            }

            document.AddParagraph("Source only");
            document.Save();

            copyStream.Position = 0;
            using WordDocument reopenedCopy = WordDocument.Load(copyStream, readOnly: true);
            Assert.Single(reopenedCopy.Paragraphs);
            Assert.Equal("Shared", reopenedCopy.Paragraphs[0].Text);

            source.Position = 0;
            using WordDocument reopenedSource = WordDocument.Load(source, readOnly: true);
            Assert.Equal(new[] { "Shared", "Source only" }, reopenedSource.Paragraphs.Select(paragraph => paragraph.Text).ToArray());
        }
    }
}
