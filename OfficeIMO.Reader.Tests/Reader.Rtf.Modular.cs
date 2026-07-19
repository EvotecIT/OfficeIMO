using OfficeIMO.Reader;
using OfficeIMO.Reader.Rtf;
using OfficeIMO.Rtf;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests;

[Collection("ReaderRegistryNonParallel")]
public sealed class ReaderRtfModularTests {
    [Fact]
    public void DocumentReaderRtf_RichResult_MapsMetadataLinksFormsTablesAndImagePayloads() {
        RtfDocument document = RtfDocument.Create();
        document.Info.Title = "Rich RTF";
        document.Info.Author = "OfficeIMO";
        RtfParagraph headerParagraph = document.AddHeader().AddParagraph("Confidential ");
        headerParagraph.AddText("header portal").SetHyperlink(new Uri("https://example.test/header"));
        RtfParagraph paragraph = document.AddParagraph("Portal ");
        paragraph.AddText("open").SetHyperlink(new Uri("https://example.test/rtf"));
        RtfNote sourceNote = document.AddNote(RtfNoteKind.Footnote);
        RtfParagraph noteParagraph = sourceNote.AddParagraph("Review note ");
        noteParagraph.AddText("note portal").SetHyperlink(new Uri("https://example.test/note"));
        paragraph.AddNoteReference(sourceNote, "1");
        RtfField field = paragraph.AddField("FORMTEXT");
        field.AddText("Ada");
        field.SetFormFieldData(data => { data.Kind = RtfFormFieldKind.Text; data.Name = "Patient"; data.Protected = true; });
        RtfTable table = document.AddTable(2, 2);
        table.Rows[0].Cells[0].AddParagraph("Name");
        table.Rows[0].Cells[1].AddParagraph("Qty");
        table.Rows[1].Cells[0].AddParagraph("Bandage");
        table.Rows[1].Cells[1].AddParagraph("4");
        RtfImage image = document.AddImage(RtfImageFormat.Png, new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 });
        image.Description = "Tiny image";

        OfficeDocumentReadResult result = RtfReaderAdapter.ReadDocument(document, "rich.rtf");

        Assert.Equal("Rich RTF", result.Source.Title);
        Assert.Equal("OfficeIMO", result.Source.Author);
        Assert.Contains(result.Links, link => link.Uri == "https://example.test/rtf");
        Assert.Contains(result.Links, link =>
            link.Uri == "https://example.test/header" &&
            link.Location.BlockAnchor?.StartsWith("rtf-header-footer-", StringComparison.Ordinal) == true);
        Assert.Contains(result.Links, link =>
            link.Uri == "https://example.test/note" &&
            link.Location.BlockAnchor?.StartsWith("rtf-note-", StringComparison.Ordinal) == true);
        OfficeDocumentFormField form = Assert.Single(result.Forms);
        Assert.Equal("Patient", form.Name);
        Assert.True(form.IsReadOnly);
        Assert.Equal("Bandage", Assert.Single(result.Tables).Rows[1][0]);
        OfficeDocumentAsset asset = Assert.Single(result.Assets, item => item.Kind == "image");
        Assert.NotNull(asset.PayloadBytes);
        Assert.True(asset.PayloadHashMatches(out _));
        Assert.Contains(result.Visuals, visual => visual.Kind == "image" && visual.PayloadHash == asset.PayloadHash);
        OfficeDocumentBlock header = Assert.Single(result.Blocks, block => block.Kind == "header-footer");
        OfficeDocumentBlock note = Assert.Single(result.Blocks, block => block.Kind == "note");
        Assert.Equal(header.Id, header.Location.BlockAnchor);
        Assert.Equal(note.Id, note.Location.BlockAnchor);
        OfficeDocumentReadResult jsonResult = OfficeDocumentReadResultJson.Deserialize(
            RtfReaderAdapter.ReadDocumentJson(document, "rich.rtf"));
        Assert.Equal(ReaderInputKind.Rtf, jsonResult.Kind);
        Assert.Contains("officeimo.reader.rtf.rich-v5", result.CapabilitiesUsed);
    }

    [Fact]
    public void DocumentReaderRtf_RichTables_ApplyRowLimitToTableBlocks() {
        RtfDocument document = RtfDocument.Create();
        RtfTable table = document.AddTable(3, 1);
        table.Rows[0].Cells[0].AddParagraph("Row 1");
        table.Rows[1].Cells[0].AddParagraph("Row 2");
        table.Rows[2].Cells[0].AddParagraph("Row 3");

        OfficeDocumentReadResult result = RtfReaderAdapter.ReadDocument(
            document,
            "bounded.rtf",
            new ReaderOptions { MaxTableRows = 1 });

        ReaderTable mapped = Assert.Single(result.Tables);
        Assert.Single(mapped.Rows);
        Assert.True(mapped.Truncated);
        OfficeDocumentBlock block = Assert.Single(result.Blocks, item => item.Kind == "table");
        Assert.Contains("Row 1", block.Text, StringComparison.Ordinal);
        Assert.DoesNotContain("Row 2", block.Text, StringComparison.Ordinal);
        Assert.DoesNotContain("Row 3", block.Text, StringComparison.Ordinal);
    }

    [Fact]
    public void DocumentReaderRtf_ReadRtfDocument_EmitsParagraphChunks() {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph("Hello RTF reader.");
        document.AddParagraph("Second paragraph.");

        var chunks = RtfReaderAdapter.Read(
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

        ReaderChunk chunk = Assert.Single(RtfReaderAdapter.Read(
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
        RtfTable outerTable = document.AddTable(1, 1);
        outerTable.Rows[0].Cells[0].AddParagraph("Outer");
        RtfTable nestedTable = outerTable.Rows[0].Cells[0].AddTable(1, 1);
        RtfParagraph cellParagraph = nestedTable.Rows[0].Cells[0].AddParagraph();
        cellParagraph.AddText("Portal").SetHyperlink(new Uri("https://example.test/patient/1"));
        RtfField field = cellParagraph.AddField("FORMTEXT");
        field.AddText("Value");
        field.SetFormFieldData(data => {
            data.Kind = RtfFormFieldKind.Text;
            data.Name = "Patient";
        });

        ReaderChunk chunk = Assert.Single(RtfReaderAdapter.Read(
            document,
            sourceName: "nested.rtf",
            readerOptions: new ReaderOptions { MaxChars = 8_000 }));

        Assert.Equal(1, chunk.Diagnostics?.LinkCount);
        Assert.Equal(1, chunk.Diagnostics?.FormFieldCount);
        Assert.Contains("Portal", chunk.Markdown, StringComparison.Ordinal);
        Assert.Contains("Outer", chunk.Markdown, StringComparison.Ordinal);
        Assert.NotNull(chunk.Tables);
        Assert.Contains("Portal", Assert.Single(chunk.Tables!).Rows[0][0], StringComparison.Ordinal);
    }

    [Fact]
    public void DocumentReaderRtf_BuilderHandler_DispatchesRtfStream() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddRtfHandler().Build();
        string rtf = CreateSampleRtf("Registry RTF");
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(rtf), writable: false);

        var chunks = reader.Read(stream, "registry.rtf").ToList();

        Assert.NotEmpty(chunks);
        Assert.Equal(ReaderInputKind.Rtf, reader.DetectKind("registry.rtf"));
        Assert.Contains(chunks, chunk =>
            chunk.Kind == ReaderInputKind.Rtf &&
            string.Equals(chunk.Location.Path, "registry.rtf", StringComparison.OrdinalIgnoreCase) &&
            chunk.Text.Contains("Registry RTF", StringComparison.Ordinal));
    }

    [Fact]
    public void DocumentReaderRtf_ReadRtfStream_NonSeekable_EnforcesMaxInputBytes() {
        string rtf = CreateSampleRtf("Too much RTF content for this limit.");
        using var stream = new NonSeekableReadStream(Encoding.UTF8.GetBytes(rtf));

        var ex = Assert.Throws<IOException>(() => RtfReaderAdapter.Read(
            stream,
            sourceName: "limited.rtf",
            readerOptions: new ReaderOptions { MaxInputBytes = 16 }).ToList());

        Assert.Contains("Input exceeds MaxInputBytes", ex.Message, StringComparison.Ordinal);
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
            IncludeDiagnostics = false,
            IncludePageLocations = true
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
        Assert.True(clone.IncludePageLocations);

        clone.RtfReadOptions!.MaxDepth = 64;

        Assert.Equal(32, options.RtfReadOptions!.MaxDepth);
        Assert.Equal(64, clone.RtfReadOptions.MaxDepth);
    }

    [Fact]
    public void ReaderRtfOptions_Defaults_To_Bounded_Core_Profile() {
        var options = new ReaderRtfOptions();

        Assert.NotNull(options.RtfReadOptions?.MaxInputBytes);
        Assert.NotNull(options.RtfReadOptions?.MaxTokenCount);
        Assert.False(options.RtfReadOptions?.ReadEmbeddedObjects);
        Assert.False(options.RtfReadOptions?.ReadFileReferences);
    }

    [Fact]
    public void DocumentReaderRtf_Reports_Object_And_Shape_Text_Fallback() {
        RtfDocument document = RtfDocument.Create();
        RtfObject rtfObject = document.AddObject(RtfObjectKind.Embedded, new byte[] { 1, 2 });
        rtfObject.Result.AddText("Object text");
        document.AddShape().AddTextBoxParagraph("Shape text");
        var options = ReaderRtfOptions.CreateTrustedProfile();

        RtfConversionResult<IReadOnlyList<ReaderChunk>> conversion = RtfReaderAdapter.ReadResult(document, rtfOptions: options);
        IReadOnlyList<ReaderChunk> chunks = conversion.Value;

        Assert.Contains(chunks, chunk => chunk.Text.Contains("Object text", StringComparison.Ordinal));
        Assert.Contains(chunks, chunk => chunk.Text.Contains("Shape text", StringComparison.Ordinal));
        Assert.Contains(conversion.Report.Diagnostics, diagnostic => diagnostic.Code == "ReaderRtfObjectFlattened" && diagnostic.Action == RtfConversionAction.Flattened);
        Assert.Contains(conversion.Report.Diagnostics, diagnostic => diagnostic.Code == "ReaderRtfShapeFlattened" && diagnostic.Action == RtfConversionAction.Flattened);
        Assert.Contains(chunks[0].Warnings!, warning => warning.StartsWith("ReaderRtfObjectFlattened:", StringComparison.Ordinal));
        OfficeDocumentReadResult rich = RtfReaderAdapter.ReadDocument(document, rtfOptions: options);
        Assert.Contains(rich.Diagnostics, diagnostic =>
            diagnostic.Code == "ReaderRtfObjectFlattened" &&
            diagnostic.Category == OfficeDocumentDiagnosticCategory.Content);

        RtfDocument cleanDocument = RtfDocument.Create();
        cleanDocument.AddParagraph("Clean RTF reader operation.");
        RtfConversionResult<IReadOnlyList<ReaderChunk>> clean = RtfReaderAdapter.ReadResult(cleanDocument, rtfOptions: options);

        Assert.DoesNotContain(clean.Report.Diagnostics, diagnostic => diagnostic.Code == "ReaderRtfObjectFlattened");
        Assert.Contains(conversion.Report.Diagnostics, diagnostic => diagnostic.Code == "ReaderRtfObjectFlattened");
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
