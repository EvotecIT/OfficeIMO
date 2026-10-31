using System.Threading.Tasks;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfAttachmentAuthoringBudgetTests {
    [Theory]
    [InlineData(PdfObjectSerializationMode.Buffered)]
    [InlineData(PdfObjectSerializationMode.ForwardOnly)]
    public void CombinedAdmissionLeavesSeekableDestinationAndDocumentIntact(PdfObjectSerializationMode mode) {
        var document = OverBudgetDocument(mode);
        byte[] sentinel = { 20, 30, 40, 50, 60 };
        using var destination = new MemoryStream();
        destination.Write(sentinel, 0, sentinel.Length);
        destination.Position = 3;

        Assert.Throws<InvalidDataException>(() => document.Save(destination));
        Assert.Equal(sentinel, destination.ToArray());
        Assert.Equal(3, destination.Position);

        PdfSaveResult failure = document.TrySaveWithEvidence(destination);
        Assert.False(failure.Succeeded);
        Assert.IsType<InvalidDataException>(failure.Exception);
        Assert.Equal(0, failure.BytesWritten);
        Assert.Equal(sentinel, destination.ToArray());
        Assert.Equal(3, destination.Position);

        document.Compose(compose => compose.Settings(settings => settings.MaxTotalEmbeddedFileBytes = 6));
        PdfExtractedAttachment[] emitted = PdfAttachmentExtractor.ExtractAttachments(document.ToBytes()).ToArray();
        Assert.Equal(new[] { "catalog.bin", "structure.bin" }, emitted.Select(file => file.FileName).OrderBy(name => name));
        Assert.Equal(new byte[] { 1, 2, 3 }, emitted.Single(file => file.FileName == "catalog.bin").Bytes);
        Assert.Equal(new byte[] { 4, 5, 6 }, emitted.Single(file => file.FileName == "structure.bin").Bytes);
    }

    [Theory]
    [InlineData(PdfObjectSerializationMode.Buffered)]
    [InlineData(PdfObjectSerializationMode.ForwardOnly)]
    public void CombinedAdmissionWritesNothingToNonSeekableDestination(PdfObjectSerializationMode mode) {
        using var destination = new NonSeekableDestination();

        Assert.Throws<InvalidDataException>(() => OverBudgetDocument(mode).Save(destination));

        Assert.Empty(destination.ToArray());
    }

    [Fact]
    public async Task AsyncCombinedAdmissionPreservesSeekableBytesAndReportsFailure() {
        var document = OverBudgetDocument(PdfObjectSerializationMode.ForwardOnly);
        byte[] sentinel = { 5, 4, 3, 2, 1 };
        using var destination = new MemoryStream();
        destination.Write(sentinel, 0, sentinel.Length);
        destination.Position = 2;

        await Assert.ThrowsAsync<InvalidDataException>(() => document.SaveAsync(destination));
        Assert.Equal(sentinel, destination.ToArray());
        Assert.Equal(2, destination.Position);

        PdfSaveResult failure = await document.TrySaveWithEvidenceAsync(destination);
        Assert.False(failure.Succeeded);
        Assert.IsType<InvalidDataException>(failure.Exception);
        Assert.Equal(sentinel, destination.ToArray());
        Assert.Equal(2, destination.Position);
    }

    [Fact]
    public async Task CombinedAdmissionPreservesAnExistingFileForSyncAndAsyncSave() {
        string path = Path.Combine(Path.GetTempPath(), "officeimo-pdf-attachment-admission-" + Guid.NewGuid().ToString("N") + ".pdf");
        byte[] sentinel = { 6, 7, 8, 9 };
        System.IO.File.WriteAllBytes(path, sentinel);
        try {
            var document = OverBudgetDocument(PdfObjectSerializationMode.Buffered);
            Assert.Throws<InvalidDataException>(() => document.Save(path));
            Assert.Equal(sentinel, System.IO.File.ReadAllBytes(path));

            await Assert.ThrowsAsync<InvalidDataException>(() => document.SaveAsync(path));
            Assert.Equal(sentinel, System.IO.File.ReadAllBytes(path));
        } finally {
            System.IO.File.Delete(path);
        }
    }

    [Fact]
    public void DocumentPerFileLimitAlsoAppliesToPermissiveStructureAdmission() {
        var options = new PdfCreateOptions {
            TaggedStructureMode = PdfTaggedStructureMode.CatalogMarkers,
            MaxEmbeddedFileBytes = 2,
            MaxTotalEmbeddedFileBytes = 10
        };
        var document = PdfDocument.Create(options).Canvas(canvas => canvas.Structure(
            PdfCanvasStructureRole.Formula, nested => nested.Text("x", global::OfficeIMO.OfficeLength.Points(20), global::OfficeIMO.OfficeLength.Points(20), global::OfficeIMO.OfficeLength.Points(40), global::OfficeIMO.OfficeLength.Points(20)),
            Association("larger.bin", new byte[] { 1, 2, 3 })));

        Assert.Throws<InvalidDataException>(() => document.ToBytes());

        document.Compose(compose => compose.Settings(settings => settings.MaxEmbeddedFileBytes = 3));
        Assert.Equal(new byte[] { 1, 2, 3 }, Assert.Single(PdfAttachmentExtractor.ExtractAttachments(document.ToBytes())).Bytes);
    }

    [Fact]
    public void CombinedBudgetCountsIdenticalStructureFilesOnceAcrossPages() {
        var options = new PdfCreateOptions {
            TaggedStructureMode = PdfTaggedStructureMode.CatalogMarkers,
            MaxEmbeddedFileBytes = 3,
            MaxTotalEmbeddedFileBytes = 6
        }.AddEmbeddedFile("catalog.bin", new byte[] { 1, 2, 3 }, "application/octet-stream");
        var association = Association("shared.bin", new byte[] { 4, 5, 6 });
        var document = PdfDocument.Create(compose => compose.Content(content => content
            .Canvas(canvas => canvas.Structure(PdfCanvasStructureRole.Formula, nested => nested.Text("x", global::OfficeIMO.OfficeLength.Points(20), global::OfficeIMO.OfficeLength.Points(20), global::OfficeIMO.OfficeLength.Points(40), global::OfficeIMO.OfficeLength.Points(20)), association))
            .PageBreak()
            .Canvas(canvas => canvas.Structure(PdfCanvasStructureRole.Formula, nested => nested.Text("y", global::OfficeIMO.OfficeLength.Points(20), global::OfficeIMO.OfficeLength.Points(20), global::OfficeIMO.OfficeLength.Points(40), global::OfficeIMO.OfficeLength.Points(20)), association))), options);

        byte[] bytes = document.ToBytes();
        Assert.Equal(2, PdfInspector.Inspect(bytes).PageCount);
        Assert.Equal(2, PdfAttachmentExtractor.ExtractAttachments(bytes).Count);

        document.Canvas(canvas => canvas.Structure(PdfCanvasStructureRole.Formula,
            nested => nested.Text("z", global::OfficeIMO.OfficeLength.Points(20), global::OfficeIMO.OfficeLength.Points(60), global::OfficeIMO.OfficeLength.Points(40), global::OfficeIMO.OfficeLength.Points(20)), Association("distinct.bin", new byte[] { 7 })));
        Assert.Throws<InvalidDataException>(() => document.ToBytes());

        document.Compose(compose => compose.Settings(settings => settings.MaxTotalEmbeddedFileBytes = 7));
        Assert.Equal(3, PdfAttachmentExtractor.ExtractAttachments(document.ToBytes()).Count);
    }

    [Fact]
    public void UnemittedStructureAttachmentsDoNotConsumeTheDocumentBudget() {
        var options = new PdfCreateOptions {
            TaggedStructureMode = PdfTaggedStructureMode.None,
            MaxEmbeddedFileBytes = 1,
            MaxTotalEmbeddedFileBytes = 1
        };
        var document = PdfDocument.Create(options).Canvas(canvas => canvas.Structure(
            PdfCanvasStructureRole.Formula, nested => nested.Text("x", global::OfficeIMO.OfficeLength.Points(20), global::OfficeIMO.OfficeLength.Points(20), global::OfficeIMO.OfficeLength.Points(40), global::OfficeIMO.OfficeLength.Points(20)),
            Association("unemitted.bin", new byte[] { 1, 2, 3 })));

        Assert.Empty(PdfAttachmentExtractor.ExtractAttachments(document.ToBytes()));
    }

    private static PdfDocument OverBudgetDocument(PdfObjectSerializationMode mode) =>
        PdfDocument.Create(new PdfCreateOptions {
            TaggedStructureMode = PdfTaggedStructureMode.CatalogMarkers,
            ObjectSerializationMode = mode,
            MaxEmbeddedFileBytes = 3,
            MaxTotalEmbeddedFileBytes = 5
        }.AddEmbeddedFile("catalog.bin", new byte[] { 1, 2, 3 }, "application/octet-stream"))
        .Canvas(canvas => canvas.Structure(PdfCanvasStructureRole.Formula,
            nested => nested.Text("x", global::OfficeIMO.OfficeLength.Points(20), global::OfficeIMO.OfficeLength.Points(20), global::OfficeIMO.OfficeLength.Points(40), global::OfficeIMO.OfficeLength.Points(20)), Association("structure.bin", new byte[] { 4, 5, 6 })));

    private sealed class NonSeekableDestination : Stream {
        private readonly MemoryStream _bytes = new();
        public override bool CanRead => false;
        public override bool CanSeek => false;
        public override bool CanWrite => true;
        public override long Length => throw new NotSupportedException();
        public override long Position {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }
        public override void Flush() => _bytes.Flush();
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => _bytes.Write(buffer, offset, count);
        internal byte[] ToArray() => _bytes.ToArray();
        protected override void Dispose(bool disposing) {
            if (disposing) _bytes.Dispose();
            base.Dispose(disposing);
        }
    }
}
