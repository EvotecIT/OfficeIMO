using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Html;
using System.Threading;

namespace OfficeIMO.Chm.Tests;

public sealed class MarkdownCancellationTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CancellationDuringHtmlProjectionStopsFurtherCallbacks(bool model) {
        ChmDocument book = ChmDocument.Load(ChmFixture.Book());
        using var cancelled = new CancellationTokenSource();
        int calls = 0;
        var options = new HtmlMarkdownExportOptions();
        options.ElementFilters.Add(_ => { calls++; cancelled.Cancel(); return false; });
        Assert.ThrowsAny<OperationCanceledException>(() => {
            if (model) book.ConvertToMarkdownDocumentWithEvidence(markdownOptions: options, cancellationToken: cancelled.Token);
            else book.ConvertToMarkdownWithEvidence(markdownOptions: options, cancellationToken: cancelled.Token);
        });
        Assert.Equal(1, calls);
    }

    [Fact]
    public void CancellationDuringMarkdownSerializationCannotReturnSuccess() {
        ChmDocument book = ChmDocument.Load(ChmFixture.Book());
        using var cancelled = new CancellationTokenSource();
        int calls = 0;
        var options = new HtmlMarkdownExportOptions { MarkdownWriteOptions = new MarkdownSaveOptions() };
        options.MarkdownWriteOptions.BlockRenderExtensions.Add(MarkdownBlockMarkdownRenderExtension.CreateContextual(
            "cancel-serialization", typeof(ParagraphBlock), (_, context) => {
                Assert.Equal(cancelled.Token, context.CancellationToken);
                calls++; cancelled.Cancel(); return "Cancelled paragraph";
            }));
        Assert.ThrowsAny<OperationCanceledException>(() => book.ConvertToMarkdownWithEvidence(markdownOptions: options, cancellationToken: cancelled.Token));
        Assert.Equal(1, calls);
    }
}
