using System;
using System.IO;
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
        public void Create_NonSeekableAssociatedStreamIsRejected() {
            using var stream = new NonSeekableWriteStream();

            ArgumentException exception = Assert.Throws<ArgumentException>(() => VisioDocument.Create(stream));

            Assert.Equal("stream", exception.ParamName);
            Assert.Contains("support seeking", exception.Message, StringComparison.OrdinalIgnoreCase);
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
