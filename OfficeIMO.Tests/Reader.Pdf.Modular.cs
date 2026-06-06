using OfficeIMO.Pdf;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Pdf;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderPdfModularTests {
    [Fact]
    public void DocumentReaderPdf_ReadPdfStream_EmitsPageAwareChunks() {
        byte[] pdf = BuildTwoPagePdf();
        using var stream = new MemoryStream(pdf, writable: false);

        var chunks = DocumentReaderPdfExtensions.ReadPdf(
            stream,
            sourceName: " sample.pdf ",
            readerOptions: new ReaderOptions { MaxChars = 8_000, ComputeHashes = true }).ToList();

        Assert.Equal(2, chunks.Count);
        Assert.All(chunks, c => {
            Assert.Equal(ReaderInputKind.Pdf, c.Kind);
            Assert.Equal("sample.pdf", c.Location.Path);
            Assert.False(string.IsNullOrWhiteSpace(c.SourceId));
            Assert.False(string.IsNullOrWhiteSpace(c.SourceHash));
            Assert.False(string.IsNullOrWhiteSpace(c.ChunkHash));
            Assert.True(c.TokenEstimate.HasValue && c.TokenEstimate.Value >= 1);
            Assert.Equal(pdf.Length, c.SourceLengthBytes);
            Assert.Null(c.SourceLastWriteUtc);
        });
        Assert.Contains(chunks, c => c.Location.Page == 1 && (c.Markdown ?? c.Text).Contains("Reader PDF page one", StringComparison.Ordinal));
        Assert.Contains(chunks, c => c.Location.Page == 2 && (c.Markdown ?? c.Text).Contains("Reader PDF page two", StringComparison.Ordinal));
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfStream_CanSelectPageRanges() {
        byte[] pdf = BuildTwoPagePdf();
        using var stream = new MemoryStream(pdf, writable: false);

        var chunks = DocumentReaderPdfExtensions.ReadPdf(
            stream,
            sourceName: "ranges.pdf",
            pdfOptions: new ReaderPdfOptions {
                PageRanges = new[] { PdfPageRange.From(2, 2) }
            }).ToList();

        var chunk = Assert.Single(chunks);
        Assert.Equal(2, chunk.Location.Page);
        Assert.Contains("Reader PDF page two", chunk.Markdown ?? chunk.Text, StringComparison.Ordinal);
        Assert.DoesNotContain("Reader PDF page one", chunk.Markdown ?? chunk.Text, StringComparison.Ordinal);
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfStream_UsesCurrentSeekableStreamPosition() {
        byte[] pdf = BuildTwoPagePdf();
        byte[] prefix = Encoding.ASCII.GetBytes("not-a-pdf-prefix");
        using var stream = new MemoryStream(prefix.Concat(pdf).ToArray(), writable: false);
        stream.Position = prefix.Length;

        var chunks = DocumentReaderPdfExtensions.ReadPdf(stream, sourceName: "embedded.pdf").ToList();

        Assert.NotEmpty(chunks);
        Assert.Contains(chunks, c => (c.Markdown ?? c.Text).Contains("Reader PDF page one", StringComparison.Ordinal));
        Assert.Contains(chunks, c => (c.Markdown ?? c.Text).Contains("Reader PDF page two", StringComparison.Ordinal));
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfStream_MaxInputBytesUsesCurrentSeekableStreamPosition() {
        byte[] pdf = BuildTwoPagePdf();
        byte[] prefix = Encoding.ASCII.GetBytes("not-a-pdf-prefix-that-is-longer-than-the-limit");
        using var stream = new MemoryStream(prefix.Concat(pdf).ToArray(), writable: false);
        stream.Position = prefix.Length;

        var chunks = DocumentReaderPdfExtensions.ReadPdf(
            stream,
            sourceName: "embedded-limited.pdf",
            readerOptions: new ReaderOptions { MaxInputBytes = pdf.Length }).ToList();

        Assert.NotEmpty(chunks);
        Assert.All(chunks, c => Assert.Equal(pdf.Length, c.SourceLengthBytes));
        Assert.Contains(chunks, c => (c.Markdown ?? c.Text).Contains("Reader PDF page one", StringComparison.Ordinal));
        Assert.Contains(chunks, c => (c.Markdown ?? c.Text).Contains("Reader PDF page two", StringComparison.Ordinal));
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfStream_DuplicatePageRangeSelectionsEmitUniqueChunkIds() {
        byte[] pdf = BuildTwoPagePdf();
        using var stream = new MemoryStream(pdf, writable: false);

        var chunks = DocumentReaderPdfExtensions.ReadPdf(
            stream,
            sourceName: "duplicate-ranges.pdf",
            pdfOptions: new ReaderPdfOptions {
                PageRanges = new[] {
                    PdfPageRange.From(1, 1),
                    PdfPageRange.From(1, 1)
                }
            }).ToList();

        Assert.Equal(2, chunks.Count);
        Assert.All(chunks, chunk => Assert.Equal(1, chunk.Location.Page));
        Assert.Equal(chunks.Count, chunks.Select(chunk => chunk.Id).Distinct(StringComparer.Ordinal).Count());
        Assert.Equal(chunks.Count, chunks.Select(chunk => chunk.Location.BlockAnchor).Distinct(StringComparer.Ordinal).Count());
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfStream_ExposesDetectedTables() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                PageWidth = 420,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Paragraph(p => p.Text("Reader PDF table marker."))
            .Table(new[] {
                new[] { "Code", "Name", "Qty" },
                new[] { "A-100", "Alpha", "2" },
                new[] { "B-200", "Beta", "14" }
            }, style: new PdfTableStyle {
                ColumnWidthPoints = new List<double?> { 70, 170, 60 },
                HeaderRowCount = 1,
                CellPaddingX = 6,
                CellPaddingY = 4
            })
            .ToBytes();
        using var stream = new MemoryStream(pdf, writable: false);

        var chunk = Assert.Single(DocumentReaderPdfExtensions.ReadPdf(
            stream,
            sourceName: "table.pdf",
            pdfOptions: new ReaderPdfOptions {
                LayoutOptions = new PdfTextLayoutOptions {
                    ForceSingleColumn = true
                }
            },
            readerOptions: new ReaderOptions { MaxChars = 8_000 }).ToList());

        Assert.NotNull(chunk.Tables);
        var table = Assert.Single(chunk.Tables!);
        Assert.Equal(3, table.TotalRowCount);
        Assert.False(table.Truncated);
        Assert.Contains(table.Rows, r => string.Join("|", r).Contains("A-100|Alpha|2", StringComparison.Ordinal));
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfStream_SplitsByMaxChars() {
        string longText = string.Join(" ", Enumerable.Repeat("Reader PDF split marker", 40));
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                PageWidth = 420,
                PageHeight = 620,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Paragraph(p => p.Text(longText))
            .ToBytes();
        using var stream = new MemoryStream(pdf, writable: false);

        var chunks = DocumentReaderPdfExtensions.ReadPdf(
            stream,
            sourceName: "split.pdf",
            readerOptions: new ReaderOptions { MaxChars = 96 }).ToList();

        Assert.True(chunks.Count > 1);
        Assert.All(chunks, c => Assert.Equal(ReaderInputKind.Pdf, c.Kind));
        Assert.Contains(chunks, c =>
            c.Warnings != null &&
            c.Warnings.Any(w => w.Contains("split due to MaxChars", StringComparison.OrdinalIgnoreCase)));
    }

    [Fact]
    public void DocumentReaderPdf_Registration_DispatchesPdfStream() {
        try {
            DocumentReaderPdfRegistrationExtensions.RegisterPdfHandler();

            byte[] pdf = BuildTwoPagePdf();
            using var stream = new MemoryStream(pdf, writable: false);
            var chunks = DocumentReader.Read(stream, "registry.pdf").ToList();

            Assert.NotEmpty(chunks);
            Assert.Contains(chunks, c =>
                c.Kind == ReaderInputKind.Pdf &&
                string.Equals(c.Location.Path, "registry.pdf", StringComparison.OrdinalIgnoreCase) &&
                (c.Markdown ?? c.Text).Contains("Reader PDF page one", StringComparison.Ordinal));
        } finally {
            DocumentReaderPdfRegistrationExtensions.UnregisterPdfHandler();
        }
    }

    [Fact]
    public void DocumentReaderPdf_BootstrapFromAssembly_RegistersPdfHandlerAndManifest() {
        try {
            DocumentReaderPdfRegistrationExtensions.UnregisterPdfHandler();

            var result = DocumentReader.BootstrapHostFromAssemblies(
                new[] { typeof(DocumentReaderPdfRegistrationExtensions).Assembly },
                new ReaderHostBootstrapOptions {
                    ReplaceExistingHandlers = true,
                    IncludeBuiltInCapabilities = true,
                    IncludeCustomCapabilities = true,
                    IndentedManifestJson = false
                });

            Assert.NotNull(result);
            Assert.Contains(result.RegisteredHandlers, handler => handler.HandlerId == DocumentReaderPdfRegistrationExtensions.HandlerId);
            Assert.Contains(result.Manifest.Handlers, handler =>
                handler.Id == DocumentReaderPdfRegistrationExtensions.HandlerId &&
                handler.Kind == ReaderInputKind.Pdf &&
                handler.Extensions.Contains(".pdf"));
            Assert.Equal(1, result.Manifest.Handlers.Count(handler =>
                string.Equals(handler.Id, DocumentReaderPdfRegistrationExtensions.HandlerId, StringComparison.Ordinal)));
            Assert.DoesNotContain(result.Manifest.Handlers, handler =>
                handler.IsBuiltIn &&
                handler.Extensions.Contains(".pdf", StringComparer.OrdinalIgnoreCase));
            Assert.Contains(DocumentReaderPdfRegistrationExtensions.HandlerId, result.ManifestJson, StringComparison.OrdinalIgnoreCase);
        } finally {
            DocumentReaderPdfRegistrationExtensions.UnregisterPdfHandler();
        }
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfStream_NonSeekable_EnforcesMaxInputBytes() {
        byte[] pdf = BuildTwoPagePdf();
        using var stream = new NonSeekableReadStream(pdf);

        var ex = Assert.Throws<IOException>(() => DocumentReaderPdfExtensions.ReadPdf(
            stream,
            sourceName: "nonseekable.pdf",
            readerOptions: new ReaderOptions { MaxInputBytes = 16 }).ToList());

        Assert.Contains("Input exceeds MaxInputBytes", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ReaderPdfOptions_CloneCopiesNestedOptionsIndependently() {
        var options = ReaderPdfOptions.CreateOfficeIMOProfile();
        options.LayoutOptions = new PdfTextLayoutOptions {
            ForceSingleColumn = true,
            MarginLeft = 12
        };
        options.PageRanges = new[] { PdfPageRange.From(1, 1) };

        var clone = options.Clone();

        Assert.NotSame(options, clone);
        Assert.NotSame(options.LayoutOptions, clone.LayoutOptions);
        Assert.NotSame(options.MarkdownOptions, clone.MarkdownOptions);
        Assert.Equal(options.LayoutOptions.MarginLeft, clone.LayoutOptions!.MarginLeft);
        Assert.Equal(options.MarkdownOptions!.IncludeLinkAnnotations, clone.MarkdownOptions!.IncludeLinkAnnotations);
        Assert.Equal(options.PageRanges.Single(), clone.PageRanges!.Single());

        clone.LayoutOptions.MarginLeft = 48;
        clone.MarkdownOptions.IncludeLinkAnnotations = false;

        Assert.Equal(12, options.LayoutOptions.MarginLeft);
        Assert.True(options.MarkdownOptions.IncludeLinkAnnotations);
    }

    private static byte[] BuildTwoPagePdf() {
        return PdfDocument.Create(new PdfOptions {
                PageWidth = 420,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .H1("Reader PDF page one")
            .Paragraph(p => p.Text("First page body."))
            .PageBreak()
            .H1("Reader PDF page two")
            .Paragraph(p => p.Text("Second page body."))
            .ToBytes();
    }
}
