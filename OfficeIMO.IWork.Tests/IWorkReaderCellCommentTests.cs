using System.Threading;
using OfficeIMO.IWork;
using OfficeIMO.Reader;
using OfficeIMO.Reader.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Reader_comment_limits_keep_original_cell_coordinates_across_flattened_headers() {
        using var package = CommentPackage(IWorkDocumentKind.Numbers);
        IWorkSourceDocument source = IWorkSourceDocument.Open(package);
        IWorkCellComment comment = source.ReadNumbers().Sheets[0].Tables[0].Cells[0].Comment!;
        var cells = new[] { (1, 1), (2, 1), (3, 1), (4, 1), (3, 2) }
            .Select(address => new IWorkTableCell(address.Item1, address.Item2,
                IWorkCellKind.Empty, null, comment: comment)).ToArray();
        var table = new IWorkTable("Headers", 4, 2, cells, headerRowCount: 2);
        var beyondBudget = new IWorkTable("Beyond", 4, 2, cells, headerRowCount: 2);
        var sheet = new IWorkNumbersSheet("Sheet", new[] { table, beyondBudget }, Array.Empty<string>());
        var numbers = new IWorkNumbersProjection(source, new[] { sheet },
            Array.Empty<IWorkDiagnostic>(), supportsEditableReconstruction: true);
        var result = new OfficeDocumentReadResult();
        var projection = new IWorkReadProjection(result, "headers.numbers",
            new ReaderOptions { MaxTableRows = 1 }, new ReaderIWorkOptions {
                MaximumTableColumns = 1, MaximumProjectedTableCells = 2
            }, CancellationToken.None);
        projection.AddNumbers(numbers);
        projection.Complete(source);

        Assert.Equal(new[] { "A1", "A3" }, result.Blocks.Where(b => b.Kind == "comment").Select(b => b.Location.A1Range));
        Assert.Equal(2, result.Metadata.Count(m => m.Category == "table.comment"));
        var omitted = Assert.Single(result.Diagnostics, d => d.Code == "IWORK_READER_TABLE_COMMENTS_OMITTED"
            && d.Attributes["tableName"] == "Headers");
        Assert.Equal("3", omitted.Attributes["omittedCommentCount"]);
        var entireTable = Assert.Single(result.Diagnostics, d => d.Code == "IWORK_READER_TABLE_COMMENTS_OMITTED"
            && d.Attributes["tableName"] == "Beyond");
        Assert.Equal("5", entireTable.Attributes["omittedCommentCount"]);
        Assert.Null(entireTable.Location!.TableIndex);
        Assert.True(Assert.Single(result.Tables).Truncated);
        Assert.Single(result.Tables[0].Rows);
    }

    [Fact]
    public void Reader_comment_chunks_retain_anchors_and_escape_literal_markup() {
        using var package = CommentPackage(IWorkDocumentKind.Numbers);
        IWorkSourceDocument source = IWorkSourceDocument.Open(package);
        IWorkCellComment native = source.ReadNumbers().Sheets[0].Tables[0].Cells[0].Comment!;
        const string text = "<b>&amp;*review* 😀 \n";
        var comment = new IWorkCellComment(text, "<author>", native.CreationDateUtc,
            native.SourceIdentity, native.SourceAuthorIdentity);
        var table = new IWorkTable("<table>", 1, 27, new[] {
            new IWorkTableCell(1, 27, IWorkCellKind.Empty, null, comment: comment)
        });
        var sheet = new IWorkNumbersSheet("Sheet", new[] { table }, Array.Empty<string>());
        var numbers = new IWorkNumbersProjection(source, new[] { sheet },
            Array.Empty<IWorkDiagnostic>(), supportsEditableReconstruction: true);
        var result = new OfficeDocumentReadResult();
        var projection = new IWorkReadProjection(result, "literal.numbers",
            new ReaderOptions { MaxChars = 1 }, new ReaderIWorkOptions(), CancellationToken.None);
        projection.AddNumbers(numbers);
        projection.Complete(source);

        ReaderChunk[] chunks = result.Chunks.Where(c => c.Location.SourceBlockKind == "table-cell-comment").ToArray();
        Assert.Equal(text, string.Concat(chunks.Select(c => c.Text)));
        Assert.All(chunks, c => { Assert.Equal("AA1", c.Location.A1Range); Assert.Equal(0, c.Location.TableIndex); });
        Assert.All(chunks.SelectMany(c => new[] { c.Text, c.Markdown }), value => {
            Assert.False(value!.Length == 1 && char.IsSurrogate(value[0]));
        });
        string markdown = string.Concat(chunks.Select(c => c.Markdown));
        string html = OfficeIMO.Markdown.MarkdownDoc.Parse(markdown).ToHtmlFragment();
        Assert.Contains("table 1, cell AA1", html);
        Assert.Contains("&lt;author&gt;", html);
        Assert.Contains("&lt;b&gt;&amp;amp;*review*", html);
        Assert.Contains("😀", System.Net.WebUtility.HtmlDecode(html));
        Assert.DoesNotContain("<b>", html);
        Assert.Equal(text, Assert.Single(result.Blocks, b => b.Kind == "comment").Text);
        Assert.Equal("<author>", Assert.Single(result.Metadata, m => m.Category == "table.comment").Attributes["author"]);
    }

    [Fact]
    public void Reader_comment_identity_metadata_counts_against_the_projection_text_budget() {
        using var package = CommentPackage(IWorkDocumentKind.Numbers);
        var limits = new IWorkReadOptions { MaximumProjectedTextCharacters = 50 };
        Assert.NotNull(IWorkSourceDocument.Open(package, limits).ReadNumbers().Sheets[0].Tables[0].Cells[0].Comment);
        package.Position = 0;
        var reader = new OfficeDocumentReaderBuilder().AddIWorkHandler(new ReaderIWorkOptions { ReadOptions = limits }).Build();
        Assert.Throws<InvalidDataException>(() => reader.ReadDocument(package, "comments.numbers"));
    }

    [Fact]
    public void Reader_comment_output_does_not_amplify_long_table_names_per_cell() {
        using var package = CommentPackage(IWorkDocumentKind.Numbers);
        IWorkSourceDocument source = IWorkSourceDocument.Open(package);
        IWorkCellComment comment = source.ReadNumbers().Sheets[0].Tables[0].Cells[0].Comment!;
        string name = new string('T', 4096);
        var table = new IWorkTable(name, 1, 64, Enumerable.Range(1, 64)
            .Select(column => new IWorkTableCell(1, column, IWorkCellKind.Empty, null, comment: comment)).ToArray());
        var numbers = new IWorkNumbersProjection(source,
            new[] { new IWorkNumbersSheet("Sheet", new[] { table }, Array.Empty<string>()) },
            Array.Empty<IWorkDiagnostic>(), supportsEditableReconstruction: true);
        var result = new OfficeDocumentReadResult { Kind = ReaderInputKind.IWork };
        var projection = new IWorkReadProjection(result, "long-title.numbers",
            new ReaderOptions(), new ReaderIWorkOptions(), CancellationToken.None);
        projection.AddNumbers(numbers);
        projection.Complete(source);
        Assert.Equal(64, result.Metadata.Count(entry => entry.Category == "table.comment"));
        Assert.Equal(name, Assert.Single(result.Tables).Title);
        Assert.True(result.ToJson().Length < name.Length * 50);
    }
}
