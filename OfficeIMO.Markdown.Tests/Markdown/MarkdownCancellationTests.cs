using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Html;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Html;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public sealed class MarkdownCancellationTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CancelledImageNameCallbackCannotPublishAnAsset(bool existingImage) {
        string root = Path.Combine(Path.GetTempPath(), "markdown-image-cancel-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string image = Path.Combine(root, "image.png");
        byte[] sentinel = [9, 8, 7];
        try {
            if (existingImage) File.WriteAllBytes(image, sentinel);
            using var cancelled = new CancellationTokenSource();
            var options = new HtmlToMarkdownOptions {
                Base64Images = HtmlBase64ImageHandling.SaveToFile,
                Base64ImageOutputDirectory = root,
                Base64ImageFileNameGenerator = (_, _) => { cancelled.Cancel(); return "image.png"; }
            };
            var document = HtmlConversionDocument.Parse("<p><img src=\"data:image/png;base64,AQID\"></p>");
            Assert.ThrowsAny<OperationCanceledException>(() => document.ToMarkdown(options, cancelled.Token));
            if (existingImage) {
                Assert.Equal(image, Assert.Single(Directory.GetFiles(root)));
                Assert.Equal(sentinel, File.ReadAllBytes(image));
            } else Assert.Empty(Directory.GetFiles(root));
        } finally { Directory.Delete(root, true); }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CancelledFallthroughBlockRendererStopsTheNextExtension(bool syntax) {
        var document = MarkdownReader.ParseWithSyntaxTree("One").Document;
        using var cancelled = new CancellationTokenSource();
        int laterCalls = 0;
        var options = new MarkdownWriteOptions();
        if (syntax) {
            options.SyntaxBlockRenderExtensions.Add(MarkdownSyntaxBlockMarkdownRenderExtension.CreateContextual(
                "later", MarkdownSyntaxKind.Paragraph, (_, _, _) => { laterCalls++; return null; }));
            options.SyntaxBlockRenderExtensions.Add(MarkdownSyntaxBlockMarkdownRenderExtension.CreateContextual(
                "cancel", MarkdownSyntaxKind.Paragraph, (_, _, _) => { cancelled.Cancel(); return null; }));
        } else {
            options.BlockRenderExtensions.Add(MarkdownBlockMarkdownRenderExtension.CreateContextual(
                "later", typeof(ParagraphBlock), (_, _) => { laterCalls++; return null; }));
            options.BlockRenderExtensions.Add(MarkdownBlockMarkdownRenderExtension.CreateContextual(
                "cancel", typeof(ParagraphBlock), (_, _) => { cancelled.Cancel(); return null; }));
        }
        Assert.ThrowsAny<OperationCanceledException>(() => document.ToMarkdown(options, cancelled.Token));
        Assert.Equal(0, laterCalls);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CancelledFallthroughInlineRendererStopsTheNextExtension(bool syntax) {
        var document = MarkdownReader.ParseWithSyntaxTree("[One](https://example.test/)").Document;
        using var cancelled = new CancellationTokenSource();
        int laterCalls = 0;
        var options = new MarkdownWriteOptions();
        if (syntax) {
            options.SyntaxInlineRenderExtensions.Add(MarkdownSyntaxInlineMarkdownRenderExtension.CreateContextual(
                "later", MarkdownSyntaxKind.InlineLink, (_, _, _) => { laterCalls++; return null; }));
            options.SyntaxInlineRenderExtensions.Add(MarkdownSyntaxInlineMarkdownRenderExtension.CreateContextual(
                "cancel", MarkdownSyntaxKind.InlineLink, (_, _, _) => { cancelled.Cancel(); return null; }));
        } else {
            options.InlineRenderExtensions.Add(MarkdownInlineMarkdownRenderExtension.CreateContextual(
                "later", typeof(LinkInline), (_, _) => { laterCalls++; return null; }));
            options.InlineRenderExtensions.Add(MarkdownInlineMarkdownRenderExtension.CreateContextual(
                "cancel", typeof(LinkInline), (_, _) => { cancelled.Cancel(); return null; }));
        }
        Assert.ThrowsAny<OperationCanceledException>(() => document.ToMarkdown(options, cancelled.Token));
        Assert.Equal(0, laterCalls);
    }

    [Fact]
    public void CancelledFallthroughHtmlInlineConverterStopsTheNextPlugin() {
        var document = HtmlConversionDocument.Parse("<p><span>One</span></p>");
        using var cancelled = new CancellationTokenSource();
        int laterCalls = 0;
        var options = new HtmlToMarkdownOptions();
        options.InlineElementConverters.Add(new HtmlInlineElementConverter("cancel", "Cancel", _ => { cancelled.Cancel(); return null; }));
        options.InlineElementConverters.Add(new HtmlInlineElementConverter("later", "Later", _ => { laterCalls++; return null; }));
        Assert.ThrowsAny<OperationCanceledException>(() => document.ToMarkdownDocumentResult(options, cancelled.Token));
        Assert.Equal(0, laterCalls);
    }

    [Fact]
    public void SharedHtmlConversionStopsWhenAnElementFilterCancels() {
        var document = HtmlConversionDocument.Parse("<p>One</p><p>Two</p>");
        using var cancelled = new CancellationTokenSource();
        int calls = 0;
        var options = new HtmlToMarkdownOptions();
        options.ElementFilters.Add(_ => { calls++; cancelled.Cancel(); return false; });
        Assert.ThrowsAny<OperationCanceledException>(() => document.ToMarkdownDocumentResult(options, cancelled.Token));
        Assert.Equal(1, calls);
        Assert.Contains("One", document.ToMarkdown());
    }

    [Fact]
    public void CancelledWriterStopsAfterTheFirstCallbackAndRestoresItsContext() {
        var document = MarkdownDoc.Create().P("One").P("Two");
        using var cancelled = new CancellationTokenSource();
        int calls = 0;
        var options = new MarkdownWriteOptions();
        options.BlockRenderExtensions.Add(MarkdownBlockMarkdownRenderExtension.CreateContextual(
            "cancel-block", typeof(ParagraphBlock), (_, context) => {
                Assert.Equal(cancelled.Token, context.CancellationToken);
                calls++; cancelled.Cancel(); return "Cancelled";
            }));
        Assert.ThrowsAny<OperationCanceledException>(() => document.ToMarkdown(options, cancelled.Token));
        Assert.Equal(1, calls);
        Assert.Contains("One", document.ToMarkdown());
        Assert.Contains("Two", document.ToMarkdown());
    }

    [Fact]
    public void CancellationInsideAnInlineRendererStopsTheRemainingInlines() {
        var document = MarkdownDoc.Create().P(paragraph => paragraph.Bold("One").Bold("Two"));
        using var cancelled = new CancellationTokenSource();
        int calls = 0;
        var options = new MarkdownWriteOptions();
        options.InlineRenderExtensions.Add(MarkdownInlineMarkdownRenderExtension.CreateContextual(
            "cancel-inline", typeof(BoldInline), (_, context) => {
                Assert.Equal(cancelled.Token, context.CancellationToken);
                calls++; cancelled.Cancel(); return "Cancelled";
            }));
        Assert.ThrowsAny<OperationCanceledException>(() => document.ToMarkdown(options, cancelled.Token));
        Assert.Equal(1, calls);
    }

    [Fact]
    public async Task CancellationWhileSerializingAnAsyncSavePreservesTheDestination() {
        string path = Path.Combine(Path.GetTempPath(), "markdown-cancel-" + Guid.NewGuid().ToString("N") + ".md");
        File.WriteAllText(path, "Existing content");
        try {
            using var cancelled = new CancellationTokenSource();
            var options = new MarkdownWriteOptions();
            options.BlockRenderExtensions.Add(MarkdownBlockMarkdownRenderExtension.CreateContextual(
                "cancel-save", typeof(ParagraphBlock), (_, _) => { cancelled.Cancel(); return "Cancelled"; }));
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => MarkdownDoc.Create().P("New content").SaveAsync(path, options, cancellationToken: cancelled.Token));
            Assert.Equal("Existing content", File.ReadAllText(path));
        } finally { File.Delete(path); }
    }
}
