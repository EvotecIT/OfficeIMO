using OfficeIMO.Html;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;

namespace OfficeIMO.Tests;

public sealed class OfficeConversionOutputCancellationTests {
    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void SynchronousAdapterSavesObserveCancellationDuringCallerStreamWrites(int route) {
        using var cancellation = new CancellationTokenSource();
        using var destination = new CancellingStream(cancellation);
        using var word = WordDocument.Create();
        word.AddParagraph("Caller-owned output cancellation");
        using var markdown = MarkdownDocument.Parse("Caller-owned output cancellation");
        var html = HtmlConversionDocument.Parse("<p>Caller-owned output cancellation</p>");

        Assert.ThrowsAny<OperationCanceledException>(() => {
            switch (route) {
                case 0: word.SaveAsMarkdown(destination, cancellationToken: cancellation.Token); break;
                case 1: markdown.SaveAsWord(destination, cancellationToken: cancellation.Token); break;
                case 2: html.SaveAsMarkdown(destination, cancellationToken: cancellation.Token); break;
            }
        });
        Assert.True(destination.OutputBegan);
        Assert.True(destination.CanWrite);
    }

    private sealed class CancellingStream(CancellationTokenSource cancellation) : MemoryStream {
        internal bool OutputBegan { get; private set; }
        public override void Write(byte[] buffer, int offset, int count) {
            base.Write(buffer, offset, count);
            OutputBegan = true;
            cancellation.Cancel();
        }
    }
}
