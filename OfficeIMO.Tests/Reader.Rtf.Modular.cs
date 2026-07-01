using OfficeIMO.Reader;
using OfficeIMO.Reader.Rtf;
using OfficeIMO.Rtf;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests;

[Collection("ReaderRegistryNonParallel")]
public sealed class ReaderRtfModularTests {
    [Fact]
    public void DocumentReaderRtf_ReadRtfDocument_EmitsParagraphChunks() {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph("Hello RTF reader.");
        document.AddParagraph("Second paragraph.");

        var chunks = DocumentReaderRtfExtensions.ReadRtfDocument(
            document,
            sourceName: "inline.rtf",
            readerOptions: new ReaderOptions { MaxChars = 8_000 }).ToList();

        Assert.Equal(2, chunks.Count);
        Assert.All(chunks, chunk => {
            Assert.Equal(ReaderInputKind.Rtf, chunk.Kind);
            Assert.Equal("inline.rtf", chunk.Location.Path);
            Assert.False(string.IsNullOrWhiteSpace(chunk.SourceId));
            Assert.False(string.IsNullOrWhiteSpace(chunk.ChunkHash));
            Assert.True(chunk.TokenEstimate.HasValue);
            Assert.Equal("rtf", chunk.Diagnostics?.SourceKind);
        });
        Assert.Contains(chunks, chunk => chunk.Text.Contains("Hello RTF reader.", StringComparison.Ordinal));
        Assert.Contains(chunks, chunk => chunk.Text.Contains("Second paragraph.", StringComparison.Ordinal));
    }

    [Fact]
    public void DocumentReaderRtf_ReadRtfDocument_ExtractsTables() {
        RtfDocument document = RtfDocument.Create();
        RtfTable table = document.AddTable(2, 2);
        table.Rows[0].Cells[0].AddParagraph("Name");
        table.Rows[0].Cells[1].AddParagraph("Qty");
        table.Rows[1].Cells[0].AddParagraph("Bandage");
        table.Rows[1].Cells[1].AddParagraph("4");

        ReaderChunk chunk = Assert.Single(DocumentReaderRtfExtensions.ReadRtfDocument(
            document,
            sourceName: "table.rtf",
            readerOptions: new ReaderOptions { MaxChars = 8_000 }));

        Assert.Equal(ReaderInputKind.Rtf, chunk.Kind);
        Assert.NotNull(chunk.Tables);
        ReaderTable readerTable = Assert.Single(chunk.Tables!);
        Assert.Equal(2, readerTable.TotalRowCount);
        Assert.Equal("Bandage", readerTable.Rows[1][0]);
        Assert.Contains("| Column 1 | Column 2 |", chunk.Markdown, StringComparison.Ordinal);
        Assert.Equal(1, chunk.Diagnostics?.TableCount);
    }

    [Fact]
    public void DocumentReaderRtf_Diagnostics_CountsNestedLinksAndFormFields() {
        RtfDocument document = RtfDocument.Create();
        RtfTable table = document.AddTable(1, 1);
        RtfParagraph cellParagraph = table.Rows[0].Cells[0].AddParagraph();
        cellParagraph.AddText("Portal").SetHyperlink(new Uri("https://example.test/patient/1"));
        RtfField field = cellParagraph.AddField("FORMTEXT");
        field.AddText("Value");
        field.SetFormFieldData(data => {
            data.Kind = RtfFormFieldKind.Text;
            data.Name = "Patient";
        });

        ReaderChunk chunk = Assert.Single(DocumentReaderRtfExtensions.ReadRtfDocument(
            document,
            sourceName: "nested.rtf",
            readerOptions: new ReaderOptions { MaxChars = 8_000 }));

        Assert.Equal(1, chunk.Diagnostics?.LinkCount);
        Assert.Equal(1, chunk.Diagnostics?.FormFieldCount);
        Assert.Contains("Portal", chunk.Markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void DocumentReaderRtf_Registration_DispatchesRtfStream() {
        try {
            DocumentReaderRtfRegistrationExtensions.RegisterRtfHandler();
            string rtf = CreateSampleRtf("Registry RTF");
            using var stream = new MemoryStream(Encoding.UTF8.GetBytes(rtf), writable: false);

            var chunks = DocumentReader.Read(stream, "registry.rtf").ToList();

            Assert.NotEmpty(chunks);
            Assert.Equal(ReaderInputKind.Rtf, DocumentReader.DetectKind("registry.rtf"));
            Assert.Contains(chunks, chunk =>
                chunk.Kind == ReaderInputKind.Rtf &&
                string.Equals(chunk.Location.Path, "registry.rtf", StringComparison.OrdinalIgnoreCase) &&
                chunk.Text.Contains("Registry RTF", StringComparison.Ordinal));
        } finally {
            DocumentReaderRtfRegistrationExtensions.UnregisterRtfHandler();
        }
    }

    [Fact]
    public void DocumentReaderRtf_ReadRtfStream_NonSeekable_EnforcesMaxInputBytes() {
        string rtf = CreateSampleRtf("Too much RTF content for this limit.");
        using var stream = new NonSeekableReadStream(Encoding.UTF8.GetBytes(rtf));

        var ex = Assert.Throws<IOException>(() => DocumentReaderRtfExtensions.ReadRtf(
            stream,
            sourceName: "limited.rtf",
            readerOptions: new ReaderOptions { MaxInputBytes = 16 }).ToList());

        Assert.Contains("Input exceeds MaxInputBytes", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void DocumentReader_DiscoverHandlerRegistrars_FindsRtfRegistrar() {
        var registrars = DocumentReader.DiscoverHandlerRegistrars(
            typeof(DocumentReaderRtfRegistrationExtensions).Assembly).ToList();

        Assert.Contains(registrars, registrar => registrar.HandlerId == DocumentReaderRtfRegistrationExtensions.HandlerId);
    }

    [Fact]
    public void ReaderRtfOptions_CloneCopiesNestedReadOptionsIndependently() {
        var options = new ReaderRtfOptions {
            RtfReadOptions = new RtfReadOptions {
                MaxDepth = 32,
                WarnOnUnsupportedCodePages = false,
                WarnOnUnsupportedDestinations = false
            },
            ChunkByBlock = false,
            IncludeHeadersAndFooters = false,
            IncludeNotes = false,
            IncludeImagePlaceholders = false,
            IncludeDiagnostics = false
        };

        ReaderRtfOptions clone = options.Clone();

        Assert.NotSame(options, clone);
        Assert.NotSame(options.RtfReadOptions, clone.RtfReadOptions);
        Assert.Equal(32, clone.RtfReadOptions?.MaxDepth);
        Assert.False(clone.ChunkByBlock);
        Assert.False(clone.IncludeHeadersAndFooters);
        Assert.False(clone.IncludeNotes);
        Assert.False(clone.IncludeImagePlaceholders);
        Assert.False(clone.IncludeDiagnostics);

        clone.RtfReadOptions!.MaxDepth = 64;

        Assert.Equal(32, options.RtfReadOptions!.MaxDepth);
        Assert.Equal(64, clone.RtfReadOptions.MaxDepth);
    }

    private static string CreateSampleRtf(string text) {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph(text);
        return document.ToRtf();
    }

    private sealed class NonSeekableReadStream : Stream {
        private readonly MemoryStream _inner;

        public NonSeekableReadStream(byte[] bytes) {
            _inner = new MemoryStream(bytes, writable: false);
        }

        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();

        public override long Position {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        public override void Flush() {
        }

        public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);

        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

        public override void SetLength(long value) => throw new NotSupportedException();

        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        protected override void Dispose(bool disposing) {
            if (disposing) {
                _inner.Dispose();
            }

            base.Dispose(disposing);
        }
    }
}
