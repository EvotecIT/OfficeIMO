using DocumentFormat.OpenXml.Packaging;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Excel.Pdf;
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

        logical.SaveAsWord(destination, PdfWordReadOptions.CreateTablesOnly());

        using WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(destination.ToArray()), false);
        Assert.NotNull(package.MainDocumentPart);
    }

    [Fact]
    public void PowerPointTableExport_WritesToNonSeekableDestination() {
        PdfLogicalDocument logical = CreateLogicalDocument();
        using var destination = new NonSeekableWriteStream();

        logical.SaveTablesAsPowerPoint(destination);

        using PresentationDocument package = PresentationDocument.Open(new MemoryStream(destination.ToArray()), false);
        Assert.NotNull(package.PresentationPart);
    }

    [Fact]
    public async Task TableConversions_ProvideReportsAndAsyncCallerOwnedStreamWrites() {
        PdfLogicalDocument logical = CreateLogicalDocument();

        PdfWordConversionResult wordResult = logical.ToWordDocumentResult(PdfWordReadOptions.CreateTablesOnly());
        PdfExcelTableImportResult excelResult = logical.ImportTablesToExcelDocumentResult();
        PdfPowerPointTableImportResult powerPointResult = logical.ImportTablesToPowerPointPresentationResult();

        using var wordDocument = wordResult.RequireNoLoss();
        using var excelDocument = excelResult.RequireNoLoss();
        using var powerPointPresentation = powerPointResult.RequireNoLoss();
        Assert.NotEmpty(wordDocument.ToBytes());
        Assert.NotEmpty(excelDocument.ToBytes());
        Assert.NotEmpty(powerPointPresentation.ToBytes());
        Assert.False(wordResult.Report.HasLoss);
        Assert.False(excelResult.Report.HasLoss);
        Assert.False(powerPointResult.Report.HasLoss);
        Assert.True(excelResult.HasOmittedPageContent);
        Assert.True(powerPointResult.HasOmittedPageContent);
        Assert.Equal(1, excelResult.Report.SourceScope.NonTableTextBlockCount);
        Assert.Equal(1, powerPointResult.Report.SourceScope.NonTableTextBlockCount);
        Assert.Equal(0, excelResult.Report.SourceScope.DetectedTableCount);
        Assert.Equal(0, powerPointResult.Report.SourceScope.DetectedTableCount);

        using var wordStream = new MemoryStream();
        using var excelStream = new MemoryStream();
        using var powerPointStream = new MemoryStream();
        await logical.SaveAsWordAsync(wordStream, PdfWordReadOptions.CreateTablesOnly());
        await logical.SaveTablesAsExcelAsync(excelStream);
        await logical.SaveTablesAsPowerPointAsync(powerPointStream);

        wordStream.WriteByte(0);
        excelStream.WriteByte(0);
        powerPointStream.WriteByte(0);
        Assert.True(wordStream.Length > 1);
        Assert.True(excelStream.Length > 1);
        Assert.True(powerPointStream.Length > 1);
    }

    [Fact]
    public async Task TableConversionAsyncWrites_HonorPreCanceledTokens() {
        PdfLogicalDocument logical = CreateLogicalDocument();
        using var destination = new MemoryStream();
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        await Assert.ThrowsAsync<OperationCanceledException>(() =>
            logical.SaveAsWordAsync(destination, PdfWordReadOptions.CreateTablesOnly(), cancellation.Token));
        Assert.Equal(0, destination.Length);
    }

    [Fact]
    public void TableConversions_ReportOmittedVectorGraphics() {
        byte[] source = PdfDocument.Create()
            .Rectangle(
                120,
                40,
                strokeColor: PdfColor.FromRgb(0, 64, 128),
                strokeWidth: 2,
                fillColor: PdfColor.FromRgb(204, 238, 255))
            .ToBytes();
        PdfLogicalDocument logical = PdfLogicalDocument.Load(source);

        PdfExcelTableImportResult excelResult = logical.ImportTablesToExcelDocumentResult();
        PdfPowerPointTableImportResult powerPointResult = logical.ImportTablesToPowerPointPresentationResult();

        Assert.Empty(logical.TextBlocks);
        Assert.Empty(logical.Images);
        Assert.True(logical.Pages[0].VectorPrimitiveCount > 0);
        Assert.True(excelResult.HasOmittedPageContent);
        Assert.True(powerPointResult.HasOmittedPageContent);
        Assert.True(excelResult.Report.SourceScope.VectorPrimitiveCount > 0);
        Assert.Equal(
            excelResult.Report.SourceScope.VectorPrimitiveCount,
            powerPointResult.Report.SourceScope.VectorPrimitiveCount);
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
