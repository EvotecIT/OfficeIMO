using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioStreamTests {
        [Fact]
        public void Create_Save_Load_Stream_Roundtrip() {
            using var stream = new MemoryStream();
            VisioDocument document = VisioDocument.Create(stream);
            document.AddPage("StreamPage");
            document.Save();

            Assert.True(stream.Length > 0);
            stream.Position = 0;

            VisioDocument loaded = VisioDocument.Load(stream);
            Assert.Single(loaded.Pages);
            Assert.Equal("StreamPage", loaded.Pages[0].Name);
        }

        [Fact]
        public void Load_SeekableStreamReadsFromBeginningAndRestoresPosition() {
            using var stream = new MemoryStream();
            VisioDocument document = VisioDocument.Create(stream);
            document.AddPage("StreamPage");
            document.Save();
            long originalPosition = stream.Length;
            stream.Position = originalPosition;

            VisioDocument loaded = VisioDocument.Load(stream);

            Assert.Single(loaded.Pages);
            Assert.Equal("StreamPage", loaded.Pages[0].Name);
            Assert.Equal(originalPosition, stream.Position);
        }

        [Fact]
        public void Create_NonSeekableAssociatedStreamIsRejected() {
            using var stream = new NonSeekableWriteStream();

            ArgumentException exception = Assert.Throws<ArgumentException>(() => VisioDocument.Create(stream));

            Assert.Equal("stream", exception.ParamName);
            Assert.Contains("support seeking", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public async Task LoadAsync_RestoresPositionAndSaveCopyPreservesAssociation() {
            string sourcePath = Path.Combine(Path.GetTempPath(), "OfficeIMO.Visio.Source." + Guid.NewGuid().ToString("N") + ".vsdx");
            string copyPath = Path.Combine(Path.GetTempPath(), "OfficeIMO.Visio.Copy." + Guid.NewGuid().ToString("N") + ".vsdx");
            try {
                VisioDocument source = VisioDocument.Create(sourcePath);
                source.AddPage("Original");
                source.Save();

                VisioDocument document = await VisioDocument.LoadAsync(sourcePath);
                document.AddPage("Copy only");
                await document.SaveCopyAsync(copyPath);

                Assert.Equal(Path.GetFullPath(sourcePath), document.FilePath);
                Assert.Single(VisioDocument.Load(sourcePath).Pages);
                Assert.Equal(2, VisioDocument.Load(copyPath).Pages.Count);

                using var stream = document.ToStream();
                stream.Position = stream.Length;
                long originalPosition = stream.Position;
                VisioDocument loaded = await VisioDocument.LoadAsync(stream);
                Assert.Equal(originalPosition, stream.Position);
                Assert.Equal(2, loaded.Pages.Count);
                stream.ReadByte();
            } finally {
                if (File.Exists(sourcePath)) File.Delete(sourcePath);
                if (File.Exists(copyPath)) File.Delete(copyPath);
            }
        }

        [Fact]
        public async Task LoadAsync_HonorsPreCanceledTokenAndRestoresPosition() {
            VisioDocument source = VisioDocument.Create();
            source.AddPage("Canceled");
            using var stream = source.ToStream();
            stream.Position = 4;
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();

            await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
                VisioDocument.LoadAsync(stream, cancellation.Token));

            Assert.Equal(4, stream.Position);
        }

        private sealed class NonSeekableWriteStream : Stream {
            private readonly MemoryStream _inner = new();

            public override bool CanRead => false;
            public override bool CanSeek => false;
            public override bool CanWrite => true;
            public override long Length => _inner.Length;
            public override long Position {
                get => _inner.Position;
                set => throw new NotSupportedException();
            }

            public override void Flush() => _inner.Flush();
            public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
            public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
            public override void SetLength(long value) => _inner.SetLength(value);
            public override void Write(byte[] buffer, int offset, int count) => _inner.Write(buffer, offset, count);
        }
    }
}
