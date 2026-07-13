using OfficeIMO.Word;
using System.Reflection;
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
            using WordDocument reopened = WordDocument.Load(source, new WordLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
            Assert.Equal("Explicit async save", Assert.Single(reopened.Paragraphs).Text);
        }

        [Fact]
        public async Task LoadAsync_StreamSaveOnDisposeCopiesEditsBackOnAsyncDispose() {
            using var source = new MemoryStream();
            using (WordDocument created = WordDocument.Create(source)) {
                created.AddParagraph("Before");
                created.Save();
            }

            source.Position = 0;
            await using (WordDocument loaded = await WordDocument.LoadAsync(source, new WordLoadOptions {
                PersistenceMode = OfficeIMO.Drawing.DocumentPersistenceMode.SaveOnDispose
            })) {
                Assert.Single(loaded.Paragraphs).SetText("After");
            }

            Assert.True(source.CanRead);
            source.Position = 0;
            using WordDocument reopened = WordDocument.Load(source, new WordLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
            Assert.Equal("After", Assert.Single(reopened.Paragraphs).Text);
        }

        [Fact]
        public async Task LoadAsync_DisposeOwnsTheActualOpenXmlPackageStream() {
            byte[] sourceBytes;
            using (WordDocument created = WordDocument.Create()) {
                created.AddParagraph("Before");
                sourceBytes = created.ToBytes();
            }

            using var source = new MemoryStream(sourceBytes);
            using WordDocument loaded = await WordDocument.LoadAsync(source);
            FieldInfo ownedStreamField = typeof(WordDocument).GetField(
                "_ownedPackageStream",
                BindingFlags.Instance | BindingFlags.NonPublic)!;
            MemoryStream ownedPackageStream = Assert.IsType<MemoryStream>(ownedStreamField.GetValue(loaded));
            Assert.Single(loaded.Paragraphs).SetText("After");

            using var output = new MemoryStream();
            await loaded.SaveAsync(output);

            using (WordDocument packageCopy = WordDocument.Load(
                       new MemoryStream(ownedPackageStream.ToArray(), writable: false),
                       new WordLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                Assert.Equal("After", Assert.Single(packageCopy.Paragraphs).Text);
            }

            loaded.Dispose();
            Assert.Throws<ObjectDisposedException>(() => ownedPackageStream.ReadByte());
        }

        [Fact]
        public void Load_StreamExplicitPersistenceDoesNotMutateSourceUntilSave() {
            using var source = new MemoryStream();
            using (WordDocument created = WordDocument.Create(source)) {
                created.AddParagraph("Before");
                created.Save();
            }

            byte[] originalBytes = source.ToArray();
            source.Position = Math.Min(7, source.Length);
            long originalPosition = source.Position;

            using (WordDocument loaded = WordDocument.Load(source)) {
                Assert.Equal(originalPosition, source.Position);
                Assert.Single(loaded.Paragraphs).SetText("After");
            }

            Assert.Equal(originalBytes, source.ToArray());
        }

        [Fact]
        public void Load_StreamSaveOnDisposeCopiesEditsBackAndLeavesStreamOpen() {
            using var source = new MemoryStream();
            using (WordDocument created = WordDocument.Create(source)) {
                created.AddParagraph("Before");
                created.Save();
            }

            using (WordDocument loaded = WordDocument.Load(source, new WordLoadOptions {
                PersistenceMode = OfficeIMO.Drawing.DocumentPersistenceMode.SaveOnDispose
            })) {
                Assert.Single(loaded.Paragraphs).SetText("After");
            }

            Assert.True(source.CanRead);
            source.Position = 0;
            using WordDocument reopened = WordDocument.Load(source, new WordLoadOptions {
                AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly
            });
            Assert.Equal("After", Assert.Single(reopened.Paragraphs).Text);
        }

        [Fact]
        public async Task LoadAsync_NonSeekableStreamDoesNotBecomePathlessSaveTarget() {
            byte[] sourceBytes;
            using (WordDocument created = WordDocument.Create()) {
                created.AddParagraph("Buffered source");
                sourceBytes = created.ToBytes();
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
        public async Task Load_NonSeekableWritableStreamDoesNotBecomePathlessSaveTarget() {
            byte[] sourceBytes;
            using (WordDocument created = WordDocument.Create()) {
                created.AddParagraph("Buffered source");
                sourceBytes = created.ToBytes();
            }

            using var source = new NonSeekableReadWriteBuffer(sourceBytes);
            using WordDocument loaded = WordDocument.Load(source);
            loaded.AddParagraph("Unsaved edit");

            Assert.Throws<InvalidOperationException>(() => loaded.Save());
            await Assert.ThrowsAsync<InvalidOperationException>(() => loaded.SaveAsync());
            Assert.Equal(sourceBytes, source.ToArray());
        }

        [Fact]
        public void Create_NonSeekableAssociatedStreamIsRejected() {
            using var stream = new NonSeekableReadWriteBuffer(Array.Empty<byte>());

            ArgumentException exception = Assert.Throws<ArgumentException>(() => WordDocument.Create(stream));

            Assert.Equal("stream", exception.ParamName);
            Assert.Contains("support seeking", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Save_StreamIsOneTimeAndDoesNotRedirectLaterParameterlessSave() {
            using var source = new MemoryStream();
            using var document = WordDocument.Create(source);
            document.AddParagraph("Original destination");
            document.Save();

            using var oneTimeDestination = new MemoryStream();
            document.Save(oneTimeDestination);
            document.AddParagraph("Source only");
            document.Save();

            using WordDocument oneTimeCopy = WordDocument.Load(oneTimeDestination,
                new WordLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
            Assert.Single(oneTimeCopy.Paragraphs);

            using WordDocument sourceCopy = WordDocument.Load(source,
                new WordLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
            Assert.Equal(new[] { "Original destination", "Source only" },
                sourceCopy.Paragraphs.Select(paragraph => paragraph.Text).ToArray());
        }

    }
}
