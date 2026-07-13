using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfTableStreamExportContracts {
    [Fact]
    public void WordTableExport_WritesToNonSeekableDestination() {
        PdfLogicalDocument logical = CreateLogicalDocument();
        using var destination = new NonSeekableWriteStream();

        logical.SaveAsWordFromPdfTables(destination);

        using WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(destination.ToArray()), false);
        Assert.NotNull(package.MainDocumentPart);
    }

    [Fact]
    public void PowerPointTableExport_WritesToNonSeekableDestination() {
        PdfLogicalDocument logical = CreateLogicalDocument();
        using var destination = new NonSeekableWriteStream();

        logical.SaveAsPowerPointFromPdfTables(destination);

        using PresentationDocument package = PresentationDocument.Open(new MemoryStream(destination.ToArray()), false);
        Assert.NotNull(package.PresentationPart);
    }

    private static PdfLogicalDocument CreateLogicalDocument() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Non-seekable table export proof"))
            .ToBytes();
        return PdfLogicalDocument.Load(source);
    }

    private sealed class NonSeekableWriteStream : Stream {
        private readonly MemoryStream _inner = new MemoryStream();

        public override bool CanRead => false;
        public override bool CanSeek => false;
        public override bool CanWrite => true;
        public override long Length => throw new NotSupportedException();
        public override long Position {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        public byte[] ToArray() => _inner.ToArray();
        public override void Flush() => _inner.Flush();
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => _inner.Write(buffer, offset, count);

        protected override void Dispose(bool disposing) {
            if (disposing) _inner.Dispose();
            base.Dispose(disposing);
        }
    }
}
