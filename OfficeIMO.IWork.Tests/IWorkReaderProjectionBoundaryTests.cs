using System.Threading;
using OfficeIMO.IWork;
using OfficeIMO.Reader;
using OfficeIMO.Reader.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void IWork_package_copy_honors_cancellation_during_read() {
        using var cancellation = new CancellationTokenSource();
        using var input = new CancellingReadStream(File.ReadAllBytes(Fixture("nim-iwork/simple.pages")),
            cancellation);

        Assert.Throws<OperationCanceledException>(() => IWorkSourceDocument.Open(input,
            IWorkDocumentKind.Pages, null, cancellation.Token));
    }

    [Fact]
    public void Keynote_reader_retains_block_roles_geometry_links_and_source_diagnostics() {
        IWorkSourceDocument source = IWorkSourceDocument.Open(
            Fixture("nim-iwork/simple.key"), IWorkDocumentKind.Keynote);
        var textStyle = new IWorkTextStyle(null, null, null, null, null, null, null,
            null, null);
        var paragraphStyle = new IWorkParagraphStyle(null, null, null, null, null,
            null, null, null, null, null, textStyle);
        IWorkTextContent Content(string text, string? hyperlink = null) => new(
            new[] { new IWorkTextParagraph(
                text.Length == 0 ? Array.Empty<IWorkTextRun>()
                    : new[] { new IWorkTextRun(text, textStyle, hyperlink) },
                paragraphStyle, null, -1, null, IWorkParagraphBreakKind.None) },
            isComplete: true, isTextComplete: true);

        var title = new IWorkTextBox(Content("Title", "https://example.com/title"),
            new IWorkGeometry(12, 23, 123, 45, 0), null, null);
        var accessible = new IWorkTextBox(new IWorkTextContent(
            Array.Empty<IWorkTextParagraph>(), true, true), null, null,
            "Accessible box");
        var empty = new IWorkTextBox(Content(string.Empty), null, null, null);
        var linked = new IWorkTextBox(new IWorkTextContent(
            Content(string.Empty).Paragraphs.Concat(Content("Real content", "javascript:alert(1)").Paragraphs)
                .ToArray(), true, true), null, "https://example.com/shape", null);
        var image = new IWorkImageAsset("diagram.png", "Data/diagram.png", "image/png",
            new byte[] { 1 }, 1, 1, new IWorkGeometry(30, 40, 50, 60, 0),
            false, null, "[click](javascript:alert(1))");
        var table = new IWorkTable("Results", 1, 1, Array.Empty<IWorkTableCell>(),
            geometry: new IWorkGeometry(70, 80, 90, 100, 0));
        var slide = new IWorkKeynoteSlide(1, "Slide", title,
            new[] { accessible, empty, linked }, Content("Speaker"),
            new[] { image }, new[] { table },
            new[] {
                new IWorkKeynoteDrawable(title, true),
                new IWorkKeynoteDrawable(accessible, false),
                new IWorkKeynoteDrawable(empty, false),
                new IWorkKeynoteDrawable(linked, false),
                new IWorkKeynoteDrawable(image),
                new IWorkKeynoteDrawable(table)
            }, isSkipped: false);
        var diagnostic = new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
            "IWORK_TEST_PROVENANCE", "Source record needs attention.",
            "Index/Document.iwa", 42);
        var projection = new IWorkKeynoteProjection(source, new[] { slide }, null,
            new[] { diagnostic }, supportsEditableReconstruction: true);
        var result = new OfficeDocumentReadResult();
        var readerProjection = new IWorkReadProjection(result, "synthetic.key",
            new ReaderOptions(), new ReaderIWorkOptions(), CancellationToken.None);

        readerProjection.AddKeynote(projection);
        readerProjection.Complete(source);

        OfficeDocumentBlock heading = Assert.Single(result.Blocks,
            block => block.Kind == "heading");
        Assert.Equal("title", heading.Location.SourceBlockKind);
        Assert.Equal(12, heading.Region!.X);
        Assert.StartsWith("# [Title]", Assert.Single(result.Chunks,
            chunk => chunk.Location.BlockAnchor == heading.Id).Markdown);
        Assert.Contains(result.Links, link => link.Uri == "https://example.com/title"
            && link.Location.BlockAnchor == heading.Id
            && link.Location.SourceBlockKind == "title");
        Assert.Contains(result.Blocks, block => block.Text == "Accessible box"
            && block.Location.SourceBlockKind == "text-box");
        Assert.Contains(result.Blocks, block => block.Text.Length == 0
            && block.Location.SourceBlockKind == "text-box");
        Assert.Contains(result.Blocks, block => block.Text == "Speaker"
            && block.Location.SourceBlockKind == "presenter-notes");
        OfficeDocumentBlock contentBlock = Assert.Single(result.Blocks,
            block => block.Text == "Real content");
        Assert.Contains(result.Links, link => link.Uri == "https://example.com/shape"
            && link.Location.BlockAnchor == contentBlock.Id);
        Assert.Contains(result.Links, link => link.Uri == "javascript:alert(1)"
            && link.Location.BlockAnchor == contentBlock.Id);
        Assert.DoesNotContain("[Real content]", result.Markdown ?? string.Empty);
        OfficeDocumentBlock imageBlock = Assert.Single(result.Blocks,
            block => block.Kind == "image");
        Assert.Equal(30, imageBlock.Region!.X);
        Assert.DoesNotContain("[click](javascript:", result.Markdown ?? string.Empty);
        Assert.Contains("\\[click\\]", result.Markdown ?? string.Empty);
        OfficeDocumentBlock tableBlock = Assert.Single(result.Blocks,
            block => block.Kind == "table");
        Assert.Equal(70, tableBlock.Region!.X);
        OfficeDocumentDiagnostic mapped = Assert.Single(result.Diagnostics,
            item => item.Code == "IWORK_TEST_PROVENANCE");
        Assert.Equal("synthetic.key!/Index/Document.iwa", mapped.Location!.Path);
        Assert.Equal("Index/Document.iwa", mapped.Attributes["entryPath"]);
        Assert.Equal("42", mapped.Attributes["recordIdentifier"]);
    }

    private sealed class CancellingReadStream : MemoryStream {
        private readonly CancellationTokenSource _cancellation;

        internal CancellingReadStream(byte[] data, CancellationTokenSource cancellation)
            : base(data, writable: false) => _cancellation = cancellation;

        public override int Read(byte[] buffer, int offset, int count) {
            int read = base.Read(buffer, offset, count);
            if (read > 0) _cancellation.Cancel();
            return read;
        }
    }
}
