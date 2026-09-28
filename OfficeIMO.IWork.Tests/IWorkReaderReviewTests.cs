using OfficeIMO.IWork;
using OfficeIMO.Reader;
using OfficeIMO.Reader.IWork;
using System.Threading;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Reader_adapter_restores_its_input_stream_after_success_and_failure() {
        using FileStream stream = File.OpenRead(Fixture("nim-iwork/simple.pages"));
        var readerOptions = new ReaderOptions();
        var iWorkOptions = new ReaderIWorkOptions();

        Assert.Equal(ReaderInputKind.IWork, IWorkReaderAdapter.ReadDocument(stream,
            "sample.pages", readerOptions, iWorkOptions, CancellationToken.None).Kind);
        Assert.Equal(0, stream.Position);

        stream.Position = 5;
        Assert.ThrowsAny<Exception>(() => IWorkReaderAdapter.ReadDocument(stream,
            "sample.pages", readerOptions, iWorkOptions, CancellationToken.None));
        Assert.Equal(5, stream.Position);
    }

    [Fact]
    public void Reader_preserves_unicode_styles_and_unique_reused_image_assets() {
        IWorkSourceDocument source = IWorkSourceDocument.Open(
            Fixture("nim-iwork/simple.pages"), IWorkDocumentKind.Pages);
        var style = new IWorkTextStyle(null, null, null, underline: true,
            strikethrough: true, null, "Example Font", null, null);
        var paragraphStyle = new IWorkParagraphStyle(null, null, null, null, null,
            null, null, null, null, null, style);
        var paragraph = new IWorkTextParagraph(new[] {
            new IWorkTextRun("A😀B", style, null)
        }, paragraphStyle, null, -1, null, IWorkParagraphBreakKind.None);
        var body = new IWorkTextContent(new[] { paragraph }, true, true);
        byte[] imageBytes = ValidPreviewPng();
        var image = new IWorkImageAsset("reused.png", "Data/reused.png", "image/png",
            imageBytes, 1, 1, null, false, null, null);
        var pages = new IWorkPagesProjection(source, body,
            Array.Empty<IWorkPagesSection>(), Array.Empty<IWorkTextBox>(),
            new[] { image }, Array.Empty<IWorkTable>(),
            new[] { new IWorkPagesDrawable(image), new IWorkPagesDrawable(image) },
            null, Array.Empty<IWorkDiagnostic>(), true);
        var result = new OfficeDocumentReadResult();
        var projection = new IWorkReadProjection(result, "sample.pages",
            new ReaderOptions { MaxChars = 2 },
            new ReaderIWorkOptions { IncludeImagePayloads = true }, CancellationToken.None);

        projection.AddPages(pages);
        projection.Complete(source);

        ReaderChunk[] textChunks = result.Chunks.Where(chunk => chunk.Text?.Length > 0).ToArray();
        Assert.Equal("A😀B", string.Concat(textChunks.Select(chunk => chunk.Text)));
        Assert.All(textChunks, chunk => AssertValidUnicode(chunk.Text!));
        Assert.All(textChunks, chunk => AssertValidUnicode(chunk.Markdown!));
        Assert.Contains("~~", result.Markdown);
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_READER_TEXT_STYLE_PARTIAL");
        Assert.Equal(2, result.Assets.Count);
        Assert.Equal(2, result.Assets.Select(asset => asset.FileName)
            .Distinct(StringComparer.OrdinalIgnoreCase).Count());
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-reader-iwork-"
            + Guid.NewGuid().ToString("N"));
        try {
            Assert.Equal(2, result.WriteAssetsToDirectory(directory).Count);
        } finally {
            if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void Reader_splits_table_text_and_markdown_on_unicode_scalar_boundaries() {
        IWorkSourceDocument source = IWorkSourceDocument.Open(
            Fixture("nim-iwork/simple.pages"), IWorkDocumentKind.Pages);
        var plainStyle = new IWorkTextStyle(null, null, null, null, null,
            null, null, null, null);
        var alignedStyle = new IWorkParagraphStyle(null, IWorkTextAlignment.Right,
            null, null, null, null, null, null, null, null, plainStyle);
        var richCellText = new IWorkTextContent(new[] {
            new IWorkTextParagraph(new[] { new IWorkTextRun("😀", plainStyle, null) },
                alignedStyle, null, -1, null, IWorkParagraphBreakKind.None)
        }, true, true);
        var table = new IWorkTable("T", 1, 1, new[] {
            new IWorkTableCell(1, 1, IWorkCellKind.Text, "😀", richText: richCellText)
        });
        var pages = new IWorkPagesProjection(source,
            new IWorkTextContent(Array.Empty<IWorkTextParagraph>(), true, true),
            Array.Empty<IWorkPagesSection>(), Array.Empty<IWorkTextBox>(),
            Array.Empty<IWorkImageAsset>(), new[] { table },
            new[] { new IWorkPagesDrawable(table) }, null,
            Array.Empty<IWorkDiagnostic>(), true);
        var result = new OfficeDocumentReadResult();
        var projection = new IWorkReadProjection(result, "sample.pages",
            new ReaderOptions { MaxChars = 1 }, new ReaderIWorkOptions(), CancellationToken.None);

        projection.AddPages(pages);
        projection.Complete(source);

        OfficeDocumentBlock block = Assert.Single(result.Blocks);
        ReaderChunk[] parts = result.Chunks.Where(chunk =>
            chunk.Location.BlockAnchor == block.Id).ToArray();
        Assert.Equal(block.Text, string.Concat(parts.Select(part => part.Text)));
        Assert.Equal(result.Tables[0].ToMarkdownTable(),
            string.Concat(parts.Select(part => part.Markdown)));
        Assert.All(parts, part => {
            AssertValidUnicode(part.Text ?? string.Empty);
            AssertValidUnicode(part.Markdown ?? string.Empty);
        });
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_READER_TABLE_STYLE_PARTIAL");
    }

    [Fact]
    public void Reader_reports_unrepresented_paragraph_alignment_without_run_style() {
        IWorkSourceDocument source = IWorkSourceDocument.Open(
            Fixture("nim-iwork/simple.pages"), IWorkDocumentKind.Pages);
        var plainStyle = new IWorkTextStyle(null, null, null, null, null,
            null, null, null, null);
        var alignedStyle = new IWorkParagraphStyle(null, IWorkTextAlignment.Right,
            null, null, null, null, null, null, null, null, plainStyle);
        var body = new IWorkTextContent(new[] {
            new IWorkTextParagraph(new[] { new IWorkTextRun("Aligned", plainStyle, null) },
                alignedStyle, null, -1, null, IWorkParagraphBreakKind.None)
        }, true, true);
        var pages = new IWorkPagesProjection(source, body,
            Array.Empty<IWorkPagesSection>(), Array.Empty<IWorkTextBox>(),
            Array.Empty<IWorkImageAsset>(), Array.Empty<IWorkTable>(),
            Array.Empty<IWorkPagesDrawable>(), null, Array.Empty<IWorkDiagnostic>(), true);
        var result = new OfficeDocumentReadResult();
        var projection = new IWorkReadProjection(result, "sample.pages",
            new ReaderOptions(), new ReaderIWorkOptions(), CancellationToken.None);

        projection.AddPages(pages);
        projection.Complete(source);

        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_READER_TEXT_STYLE_PARTIAL");
    }

    private static void AssertValidUnicode(string value) {
        for (int index = 0; index < value.Length; index++) {
            if (char.IsHighSurrogate(value[index])) {
                Assert.True(index + 1 < value.Length && char.IsLowSurrogate(value[index + 1]));
                index++;
            } else {
                Assert.False(char.IsLowSurrogate(value[index]));
            }
        }
    }
}
