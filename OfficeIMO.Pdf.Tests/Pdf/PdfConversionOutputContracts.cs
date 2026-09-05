using System.Threading;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfConversionOutputContracts {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PdfOutputPathsShareOverwriteRewindAndOwnership(bool resultPath) {
        PdfDocument pdf = PdfDocument.Create();
        pdf.Content.Paragraph(p => p.Text("Output contract"));
        using var destination = new MemoryStream();
        destination.Write(new byte[50000], 0, 50000);
        destination.Position = 37;

        PdfSaveResult result = resultPath
            ? new PdfDocumentConversionResult(pdf, new PdfConversionReport()).Save(destination)
            : pdf.Save(destination);

        Assert.True(result.RequireSuccess().Succeeded);
        Assert.Equal(0, destination.Position);
        Assert.True(destination.Length < 50000);
        Assert.Contains("Output contract", PdfDocument.Load(destination.ToArray()).Read().Text);
        Assert.True(destination.CanWrite);
    }

    [Fact]
    public void WordPdfOutputUsesTheSameStreamPolicy() {
        using var word = WordDocument.Create();
        word.AddParagraph("Word output contract");
        using var destination = new MemoryStream();
        destination.Write(new byte[50000], 0, 50000);
        destination.Position = 37;
        word.SaveAsPdf(destination).RequireSuccess();
        Assert.Equal(0, destination.Position);
        Assert.True(destination.Length < 50000);
        Assert.True(destination.CanWrite);
        Assert.Contains("Word output contract", PdfDocument.Load(destination.ToArray()).Read().Text);
    }

    [Fact]
    public void SharedOutputGuardRejectsFailureAndReverseSaveRetainsTypedEvidence() {
        var failure = OfficeOutputResult<PdfConversionReport>.FromFailure(null, new IOException("Write failed"));
        Assert.False(((IOfficeResult)failure).Succeeded);
        Assert.Throws<InvalidOperationException>(() => failure.RequireNoLoss());

        PdfDocument pdf = PdfDocument.Create();
        pdf.Content.Paragraph(p => p.Text("HTML output"));
        using var destination = new MemoryStream();
        OfficeOutputResult<PdfConversionReport> result = PdfDocument.Load(pdf.ToBytes()).SaveAsHtml(destination);
        Assert.True(((IOfficeOutputResult)result).Succeeded);
        Assert.Null(result.OutputPath);
        Assert.NotNull(result.Report);
        Assert.Contains("HTML output", System.Text.Encoding.UTF8.GetString(destination.ToArray()));
        Assert.True(destination.CanWrite);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CancellationDuringLayoutPropagatesThroughByteAndResultSaves(bool saveResult) {
        using var cancellation = new CancellationTokenSource();
        bool enteredLayout = false;
        PdfDocument pdf = PdfDocument.Create();
        pdf.Content.Deferred(_ => {
            enteredLayout = true;
            cancellation.Cancel();
            return content => content.Paragraph(p => p.Text("Canceled during layout"));
        });
        var conversion = new PdfDocumentConversionResult(pdf, new PdfConversionReport());
        using var destination = new MemoryStream();
        if (saveResult) {
            Assert.Throws<OperationCanceledException>(() => conversion.SaveResult(destination, cancellation.Token));
        } else {
            Assert.Throws<OperationCanceledException>(() => conversion.ToBytes(cancellation.Token));
        }
        Assert.True(enteredLayout);
        Assert.Equal(0, destination.Length);
    }

    [Fact]
    public void MarkdownCancellationReachesPdfRenderingAndDoesNotChangeReusableOptions() {
        using var cancellation = new CancellationTokenSource();
        bool enteredLayout = false;
        var options = new MarkdownToPdfOptions {
            PdfOptions = new PdfOptions {
                TextLineBreakCallback = text => {
                    enteredLayout = true;
                    cancellation.Cancel();
                    return Array.Empty<int>();
                }
            }
        };
        MarkdownDoc markdown = MarkdownDoc.Create().P(new string('x', 1000));
        Assert.Throws<OperationCanceledException>(() => markdown.ToPdfBytes(options, cancellation.Token));
        Assert.True(enteredLayout);
        Assert.NotEmpty(markdown.ToPdfBytes(options));
    }
}
