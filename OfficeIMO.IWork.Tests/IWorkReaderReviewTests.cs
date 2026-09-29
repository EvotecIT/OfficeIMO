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

    [Fact]
    public void Reader_bounds_escaped_and_linked_markdown_chunks() {
        IWorkSourceDocument source = IWorkSourceDocument.Open(
            Fixture("nim-iwork/simple.pages"), IWorkDocumentKind.Pages);
        var style = new IWorkTextStyle(null, null, null, null, null,
            null, null, null, null);
        var paragraphStyle = new IWorkParagraphStyle(null, null, null, null, null,
            null, null, null, null, null, style);
        string text = new string('`', 20);
        var paragraph = new IWorkTextParagraph(new[] {
            new IWorkTextRun(text, style, "https://example.com/" + new string('x', 100))
        }, paragraphStyle, null, -1, null, IWorkParagraphBreakKind.None);
        var pages = new IWorkPagesProjection(source,
            new IWorkTextContent(new[] { paragraph }, true, true),
            Array.Empty<IWorkPagesSection>(), Array.Empty<IWorkTextBox>(),
            Array.Empty<IWorkImageAsset>(), Array.Empty<IWorkTable>(),
            Array.Empty<IWorkPagesDrawable>(), null, Array.Empty<IWorkDiagnostic>(), true);
        var result = new OfficeDocumentReadResult();
        var projection = new IWorkReadProjection(result, "sample.pages",
            new ReaderOptions { MaxChars = 8 }, new ReaderIWorkOptions(), CancellationToken.None);

        projection.AddPages(pages);
        projection.Complete(source);

        Assert.Equal(text, string.Concat(result.Chunks.Select(chunk => chunk.Text)));
        Assert.Equal(IWorkReadProjection.RichTextMarkdown(paragraph), result.Markdown);
        Assert.All(result.Chunks, chunk => Assert.InRange(chunk.Markdown!.Length, 0, 8));
    }

    [Fact]
    public void Reader_keeps_whitespace_outside_rich_text_delimiters() {
        var plain = new IWorkTextStyle(null, null, null, null, null,
            null, null, null, null);
        var bold = new IWorkTextStyle(null, true, null, null, null,
            null, null, null, null);
        var italic = new IWorkTextStyle(null, null, true, null, null,
            null, null, null, null);
        var strike = new IWorkTextStyle(null, null, null, null, true,
            null, null, null, null);
        var paragraphStyle = new IWorkParagraphStyle(null, null, null, null, null,
            null, null, null, null, null, plain);
        var paragraph = new IWorkTextParagraph(new[] {
            new IWorkTextRun(" bold ", bold, null),
            new IWorkTextRun(" italic ", italic, null),
            new IWorkTextRun(" strike ", strike, null)
        }, paragraphStyle, null, -1, null, IWorkParagraphBreakKind.None);

        string markdown = IWorkReadProjection.RichTextMarkdown(paragraph);

        Assert.Contains(" **bold** ", markdown);
        Assert.Contains(" *italic* ", markdown);
        Assert.Contains(" ~~strike~~ ", markdown);
    }

    [Fact]
    public void Reader_rich_text_markdown_observes_cancellation_during_a_large_run() {
        var style = new IWorkTextStyle(null, null, null, null, null,
            null, null, null, null);
        var paragraphStyle = new IWorkParagraphStyle(null, null, null, null, null,
            null, null, null, null, null, style);
        var paragraph = new IWorkTextParagraph(new[] {
            new IWorkTextRun(new string('*', 4 * 1024 * 1024), style, null)
        }, paragraphStyle, null, -1, null, IWorkParagraphBreakKind.None);
        using var cancellation = new CancellationTokenSource();
        cancellation.CancelAfter(TimeSpan.FromMilliseconds(1));

        Assert.Throws<OperationCanceledException>(() =>
            IWorkReadProjection.RichTextMarkdown(paragraph, cancellation.Token));
    }

    [Fact]
    public void Reader_reports_rotations_and_merges_and_indexes_each_table_chunk() {
        IWorkSourceDocument source = IWorkSourceDocument.Open(
            Fixture("nim-iwork/simple.pages"), IWorkDocumentKind.Pages);
        var rotation = new IWorkGeometry(10, 20, 100, 40, 15);
        var emptyText = new IWorkTextContent(Array.Empty<IWorkTextParagraph>(), true, true);
        var textBox = new IWorkTextBox(emptyText, rotation, null, "Rotated text box");
        var image = new IWorkImageAsset("image.png", "Data/image.png", "image/png",
            ValidPreviewPng(), 1, 1, rotation, false, null, null);
        var merged = new IWorkTable("Merged", 1, 2,
            new[] { new IWorkTableCell(1, 1, IWorkCellKind.Text, "Value") },
            mergedRanges: new[] { new IWorkTableMergeRange(1, 1, 1, 2) },
            geometry: rotation);
        var second = new IWorkTable("Second", 1, 1,
            new[] { new IWorkTableCell(1, 1, IWorkCellKind.Text, "Other") });
        var pages = new IWorkPagesProjection(source, emptyText,
            Array.Empty<IWorkPagesSection>(), new[] { textBox },
            new[] { image }, new[] { merged, second },
            new IWorkPagesDrawable[] {
                new(textBox), new(image), new(merged), new(second)
            }, null, Array.Empty<IWorkDiagnostic>(), true);
        var result = new OfficeDocumentReadResult();
        var projection = new IWorkReadProjection(result, "sample.pages",
            new ReaderOptions { MaxChars = 4 }, new ReaderIWorkOptions(), CancellationToken.None);

        projection.AddPages(pages);
        projection.Complete(source);

        Assert.Equal(3, result.Diagnostics.Count(diagnostic =>
            diagnostic.Code == "IWORK_READER_ROTATION_UNSUPPORTED"));
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_READER_TABLE_MERGES_UNSUPPORTED");
        OfficeDocumentBlock[] tableBlocks = result.Blocks.Where(block => block.Kind == "table").ToArray();
        Assert.Equal(2, tableBlocks.Length);
        for (int index = 0; index < tableBlocks.Length; index++) {
            Assert.Equal(index, tableBlocks[index].Location.TableIndex);
            ReaderChunk[] chunks = result.Chunks.Where(chunk =>
                chunk.Location.BlockAnchor == tableBlocks[index].Id).ToArray();
            Assert.NotEmpty(chunks);
            Assert.All(chunks, chunk => Assert.Equal(index, chunk.Location.TableIndex));
        }
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Reader_reports_formula_cache_styles_it_cannot_project(IWorkDocumentKind kind) {
        using MemoryStream package = CreateFormulaTableWithIncompleteRichCacheStyle(kind);

        OfficeDocumentReadResult result = IWorkReaderAdapter.ReadDocument(package,
            kind == IWorkDocumentKind.Pages ? "sample.pages" : "sample.key",
            new ReaderOptions(), new ReaderIWorkOptions(), CancellationToken.None);

        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_READER_TABLE_STYLE_PARTIAL");
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_READER_FORMULA_CACHE");
    }

    [Theory]
    [InlineData(IWorkParagraphBreakKind.Page)]
    [InlineData(IWorkParagraphBreakKind.Section)]
    [InlineData(IWorkParagraphBreakKind.Layout)]
    public void Reader_reports_explicit_layout_breaks(IWorkParagraphBreakKind breakKind) {
        IWorkSourceDocument source = IWorkSourceDocument.Open(
            Fixture("nim-iwork/simple.pages"), IWorkDocumentKind.Pages);
        var style = new IWorkTextStyle(null, null, null, null, null,
            null, null, null, null);
        var paragraphStyle = new IWorkParagraphStyle(null, null, null, null, null,
            null, null, null, null, null, style);
        var paragraph = new IWorkTextParagraph(new[] {
            new IWorkTextRun("Before break", style, null)
        }, paragraphStyle, null, -1, null, breakKind);
        var pages = new IWorkPagesProjection(source,
            new IWorkTextContent(new[] { paragraph }, true, true),
            Array.Empty<IWorkPagesSection>(), Array.Empty<IWorkTextBox>(),
            Array.Empty<IWorkImageAsset>(), Array.Empty<IWorkTable>(),
            Array.Empty<IWorkPagesDrawable>(), null, Array.Empty<IWorkDiagnostic>(), true);
        var result = new OfficeDocumentReadResult();
        var projection = new IWorkReadProjection(result, "sample.pages",
            new ReaderOptions(), new ReaderIWorkOptions(), CancellationToken.None);

        projection.AddPages(pages);
        projection.Complete(source);

        Assert.Contains("Before break", result.Markdown);
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_READER_LAYOUT_BREAK_UNSUPPORTED");
    }

    [Fact]
    public void Reader_bounds_dense_cells_across_sparse_tables() {
        IWorkSourceDocument source = IWorkSourceDocument.Open(
            Fixture("nim-iwork/simple.numbers"), IWorkDocumentKind.Numbers);
        IWorkTable[] tables = Enumerable.Range(1, 3)
            .Select(index => new IWorkTable("Sparse " + index, 2, 2,
                Array.Empty<IWorkTableCell>()))
            .ToArray();
        var sheet = new IWorkNumbersSheet("Sheet 1", tables, Array.Empty<string>());
        var numbers = new IWorkNumbersProjection(source, new[] { sheet },
            Array.Empty<IWorkDiagnostic>(), supportsEditableReconstruction: true);
        var result = new OfficeDocumentReadResult();
        var projection = new IWorkReadProjection(result, "sample.numbers",
            new ReaderOptions(), new ReaderIWorkOptions {
                MaximumProjectedTableCells = 5
            }, CancellationToken.None);

        projection.AddNumbers(numbers);
        projection.Complete(source);

        Assert.Equal(2, result.Tables.Count);
        Assert.Equal(2, result.Tables[0].Columns.Count);
        Assert.Equal(2, result.Tables[0].Rows.Count);
        Assert.Single(result.Tables[1].Columns);
        Assert.Single(result.Tables[1].Rows);
        Assert.True(result.Tables[1].Truncated);
        Assert.Equal(5, result.Tables.Sum(table => table.Columns.Count * table.Rows.Count));
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_READER_TABLE_TRUNCATED");
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_READER_TABLE_BUDGET_EXCEEDED");
    }

    [Fact]
    public void Reader_links_retain_image_identity_and_empty_text_box_geometry() {
        IWorkSourceDocument source = IWorkSourceDocument.Open(
            Fixture("nim-iwork/simple.pages"), IWorkDocumentKind.Pages);
        var emptyText = new IWorkTextContent(Array.Empty<IWorkTextParagraph>(), true, true);
        var firstBox = new IWorkTextBox(emptyText, new IWorkGeometry(10, 20, 30, 40, 0),
            "https://example.com/box-one", null);
        var secondBox = new IWorkTextBox(emptyText, new IWorkGeometry(50, 60, 30, 40, 0),
            "https://example.com/box-two", null);
        byte[] imageBytes = ValidPreviewPng();
        var firstImage = new IWorkImageAsset("first.png", "Data/first.png", "image/png",
            imageBytes, 1, 1, new IWorkGeometry(100, 20, 30, 40, 0), false,
            "https://example.com/image-one", null);
        var secondImage = new IWorkImageAsset("second.png", "Data/second.png", "image/png",
            imageBytes, 1, 1, new IWorkGeometry(150, 20, 30, 40, 0), false,
            "https://example.com/image-two", null);
        var pages = new IWorkPagesProjection(source, emptyText,
            Array.Empty<IWorkPagesSection>(), new[] { firstBox, secondBox },
            new[] { firstImage, secondImage }, Array.Empty<IWorkTable>(),
            new IWorkPagesDrawable[] {
                new(firstBox), new(secondBox), new(firstImage), new(secondImage)
            }, null, Array.Empty<IWorkDiagnostic>(), true);
        var result = new OfficeDocumentReadResult();
        var projection = new IWorkReadProjection(result, "sample.pages",
            new ReaderOptions(), new ReaderIWorkOptions(), CancellationToken.None);

        projection.AddPages(pages);
        projection.Complete(source);

        Assert.Equal(4, result.Links.Count);
        Assert.Equal(4, result.Links.Select(link => link.Location.BlockAnchor)
            .Distinct(StringComparer.Ordinal).Count());
        Assert.Equal(new double[] { 10, 50 }, result.Links.Take(2)
            .Select(link => link.Region!.X));
        Assert.All(result.Links.Take(2), link => {
            Assert.Equal("text-box", link.Location.SourceBlockKind);
            Assert.StartsWith("iwork-shape-", link.Location.BlockAnchor);
        });
        foreach (OfficeDocumentAsset asset in result.Assets) {
            OfficeDocumentLink link = Assert.Single(result.Links,
                item => item.Location.BlockAnchor == asset.Id);
            Assert.Equal("image", link.Location.SourceBlockKind);
            Assert.Equal(asset.Region!.X, link.Region!.X);
        }
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
