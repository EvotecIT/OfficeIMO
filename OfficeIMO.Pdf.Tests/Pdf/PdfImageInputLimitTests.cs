using System;
using System.IO;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfImageInputLimitTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ImageStampBoundsNonSeekableEncodedInputBeforeBuffering(bool watermark) {
        byte[] image = PdfPngTestImages.CreateRgbPng(25, 50, 75);
        using var stream = new ChunkedNonSeekableStream(image, maximumChunkSize: 3);
        PdfDocument document = PdfDocument.Load(BuildPdf());
        var options = new PdfImageStampOptions { MaximumEncodedImageBytes = 16 };

        Assert.Throws<InvalidDataException>(() => watermark
            ? document.Stamp.ImageWatermark(stream, options)
            : document.Stamp.Image(stream, options));
        Assert.InRange(stream.BytesRead, 17, 19);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ImageStampRejectsOversizedSeekableInputBeforeReading(bool watermark) {
        using var stream = new LengthOnlyReadStream(PdfImageInput.DefaultMaximumEncodedBytes + 1L);
        PdfDocument document = PdfDocument.Load(BuildPdf());

        Assert.Throws<InvalidDataException>(() => watermark
            ? document.Stamp.ImageWatermark(stream)
            : document.Stamp.Image(stream));
        Assert.False(stream.WasRead);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ImageStampBoundsCallerBytesBeforeMutation(bool watermark) {
        byte[] image = PdfPngTestImages.CreateRgbPng(25, 50, 75);
        PdfDocument document = PdfDocument.Load(BuildPdf());
        var options = new PdfImageStampOptions { MaximumEncodedImageBytes = image.Length - 1L };

        Assert.Throws<InvalidDataException>(() => watermark
            ? document.Stamp.ImageWatermark(image, options)
            : document.Stamp.Image(image, options));
        Assert.Empty(document.Images.Extract());
    }

    [Fact]
    public void ImageStampReadsOnlyRemainingEncodedStreamBytes() {
        byte[] image = PdfPngTestImages.CreateRgbPng(25, 50, 75);
        byte[] prefixed = new byte[image.Length + 7];
        Buffer.BlockCopy(image, 0, prefixed, 7, image.Length);
        using var stream = new MemoryStream(prefixed);
        stream.Position = 7;
        var options = new PdfImageStampOptions { MaximumEncodedImageBytes = image.Length };

        PdfDocument stamped = PdfDocument.Load(BuildPdf()).Stamp.Image(stream, options);

        Assert.Single(stamped.Images.Extract());
        Assert.Equal(stream.Length, stream.Position);
    }

    private static byte[] BuildPdf() => PdfDocument.Create()
        .Paragraph(paragraph => paragraph.Text("Bounded image source"))
        .ToBytes();

    private sealed class ChunkedNonSeekableStream : Stream {
        private readonly byte[] _bytes;
        private readonly int _maximumChunkSize;
        private int _position;

        internal ChunkedNonSeekableStream(byte[] bytes, int maximumChunkSize) {
            _bytes = bytes;
            _maximumChunkSize = maximumChunkSize;
        }

        internal int BytesRead => _position;
        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position {
            get => _position;
            set => throw new NotSupportedException();
        }

        public override int Read(byte[] buffer, int offset, int count) {
            int available = _bytes.Length - _position;
            if (available <= 0) return 0;
            int read = Math.Min(Math.Min(count, _maximumChunkSize), available);
            Buffer.BlockCopy(_bytes, _position, buffer, offset, read);
            _position += read;
            return read;
        }

        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }

    private sealed class LengthOnlyReadStream : Stream {
        private long _position;

        internal LengthOnlyReadStream(long length) {
            Length = length;
        }

        internal bool WasRead { get; private set; }
        public override bool CanRead => true;
        public override bool CanSeek => true;
        public override bool CanWrite => false;
        public override long Length { get; }
        public override long Position {
            get => _position;
            set => _position = value;
        }

        public override int Read(byte[] buffer, int offset, int count) {
            WasRead = true;
            return 0;
        }

        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin) {
            _position = origin switch {
                SeekOrigin.Begin => offset,
                SeekOrigin.Current => _position + offset,
                SeekOrigin.End => Length + offset,
                _ => throw new ArgumentOutOfRangeException(nameof(origin))
            };
            return _position;
        }
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }
}
