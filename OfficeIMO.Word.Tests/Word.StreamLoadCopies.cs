using System.IO.Compression;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        private const string StreamLoadLargePart = "word/media/large.bin";

        private static byte[] CreateStreamLoadDocument(string text) {
            using WordDocument document = WordDocument.Create();
            document.AddParagraph(text);
            return document.ToBytes();
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public async Task ReadOnlySeekableLoad_ReadsCallerStreamInPlaceAndLeavesItOpenAndUnchanged(bool async) {
            byte[] original = CreateStreamLoadDocument("In place");
            using var source = new MemoryStream();
            source.Write(original, 0, original.Length);
            source.Position = 5;
            var options = new WordLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly };

            WordDocument document = async
                ? await WordDocument.LoadAsync(source, options)
                : WordDocument.Load(source, options);
            using (document) {
                Assert.Equal(5, source.Position);
                Assert.Equal("In place", Assert.Single(document.Paragraphs).Text);
                Assert.Throws<InvalidOperationException>(() => document.Save(source));
                Assert.Throws<InvalidOperationException>(() => document.Save());
            }

            Assert.True(source.CanRead);
            Assert.Equal(original, source.ToArray());
        }

        [Fact]
        public void ReadOnlyNonSeekableLoad_DoesNotNeedTheStreamAfterLoading() {
            byte[] original = CreateStreamLoadDocument("Forward only");
            WordDocument document;
            using (var source = new NonSeekableReadStream(original)) {
                document = WordDocument.Load(source, new WordLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly });
            }

            using (document) {
                Assert.Equal("Forward only", Assert.Single(document.Paragraphs).Text);
            }
        }

        [Fact]
        public void ReadWriteLoad_DoesNotShareTheCallersBuffer() {
            byte[] callerBuffer = CreateStreamLoadDocument("Detached");
            using var source = new MemoryStream(callerBuffer, writable: false);
            using WordDocument document = WordDocument.Load(source);

            Array.Clear(callerBuffer, 0, callerBuffer.Length);

            Assert.Equal("Detached", Assert.Single(document.Paragraphs).Text);
            using WordDocument reloaded = WordDocument.Load(new MemoryStream(document.ToBytes()),
                new WordLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly });
            Assert.Equal("Detached", Assert.Single(reloaded.Paragraphs).Text);
        }

        [Fact]
        public void ReadWriteLoad_SavesEditsAndKeepsUntouchedPartsIntact() {
            byte[] original = LargePartPackages.WithLargePart(CreateStreamLoadDocument("Before"), "word", 256 * 1024);
            byte[] largePart = ReadStreamLoadEntry(original, StreamLoadLargePart);
            using var source = new MemoryStream(original, writable: false);
            using WordDocument document = WordDocument.Load(source);
            Assert.Single(document.Paragraphs).SetText("After");

            using var saved = new MemoryStream();
            document.Save(saved);
            byte[] savedBytes = saved.ToArray();

            Assert.Equal(largePart, ReadStreamLoadEntry(savedBytes, StreamLoadLargePart));
            using WordDocument reloaded = WordDocument.Load(new MemoryStream(savedBytes),
                new WordLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly });
            Assert.Equal("After", Assert.Single(reloaded.Paragraphs).Text);
        }

        private static byte[] ReadStreamLoadEntry(byte[] package, string name) {
            using var archive = new ZipArchive(new MemoryStream(package, writable: false), ZipArchiveMode.Read);
            using Stream entry = archive.GetEntry(name)!.Open();
            using var copy = new MemoryStream();
            entry.CopyTo(copy);
            return copy.ToArray();
        }

#if NET
        [Fact]
        [Trait("Category", "Performance")]
        public void Load_AllocatesNoMoreCopiesThanTheAccessModeNeeds() {
            byte[] original = LargePartPackages.WithLargePart(CreateStreamLoadDocument("Large"), "word");

            long readOnly = MeasureStreamLoad(original, OfficeIMO.DocumentAccessMode.ReadOnly);
            long readWrite = MeasureStreamLoad(original, OfficeIMO.DocumentAccessMode.ReadWrite);

            Assert.True(readOnly < original.Length / 4, $"Read-only load allocated {readOnly} bytes for a {original.Length} byte file.");
            Assert.True(readWrite < original.Length * 3L / 2, $"Editable load allocated {readWrite} bytes for a {original.Length} byte file.");
        }

        private static long MeasureStreamLoad(byte[] file, OfficeIMO.DocumentAccessMode mode) {
            using var source = new MemoryStream(file, writable: false);
            GC.Collect();
            long before = GC.GetAllocatedBytesForCurrentThread();
            using WordDocument document = WordDocument.Load(source, new WordLoadOptions { AccessMode = mode });
            return GC.GetAllocatedBytesForCurrentThread() - before;
        }
#endif
    }
}
