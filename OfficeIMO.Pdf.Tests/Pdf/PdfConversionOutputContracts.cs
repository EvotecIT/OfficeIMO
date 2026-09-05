using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Html;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Pdf;
using OfficeIMO.Mhtml;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfConversionOutputContracts {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlStreamSavesObserveCancellationWhenOutputBegins(bool resultPath) {
        using var cancellation = new CancellationTokenSource();
        using var destination = new CancelWhenOutputBeginsStream(cancellation);
        HtmlConversionDocument html = HtmlConversionDocument.Parse("<p>Cancel after HTML conversion</p>");
        if (resultPath) {
            Assert.Throws<OperationCanceledException>(() => html.SaveAsPdfResult(destination, cancellationToken: cancellation.Token));
        } else {
            Assert.Throws<OperationCanceledException>(() => html.SaveAsPdf(destination, cancellationToken: cancellation.Token));
        }
        Assert.True(destination.OutputBegan);
        Assert.Equal(0, destination.Length);
    }

    [Fact]
    public void MhtmlSynchronousConversionObservesCancellationDuringHtmlPolicyEvaluation() {
        using var cancellation = new CancellationTokenSource();
        bool evaluatedPolicy = false;
        var options = new HtmlToPdfOptions();
        options.UrlPolicy.ResolvedUrlTransform = value => {
            evaluatedPolicy = true;
            cancellation.Cancel();
            return value;
        };
        var mhtml = new MhtmlDocument("<a href='https://example.test/next'>Linked text</a>");
        Assert.Throws<OperationCanceledException>(() => mhtml.ToPdfDocumentResult(options, cancellation.Token));
        Assert.True(evaluatedPolicy);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task PowerPointFileOutputsRetainTheirDestination(bool logicalSource, bool asyncOutput) {
        PdfDocument generated = PdfDocument.Create();
        generated.Content.Paragraph(p => p.Text("Presentation output"));
        PdfDocument opened = PdfDocument.Load(generated.ToBytes());
        var options = PdfToPowerPointOptions.CreateEditableTables();
        string path = Path.Combine(Path.GetTempPath(), "officeimo-output-contract-" + Guid.NewGuid().ToString("N") + ".pptx");
        try {
            OfficeOutputResult<PdfPowerPointConversionReport> result = logicalSource
                ? asyncOutput ? await opened.Read().SaveAsPowerPointAsync(path, options) : opened.Read().SaveAsPowerPoint(path, options)
                : asyncOutput ? await opened.SaveAsPowerPointAsync(path, options) : opened.SaveAsPowerPoint(path, options);
            Assert.True(result.RequireSuccess().Succeeded);
            Assert.Equal(path, result.OutputPath);
            Assert.True(new FileInfo(path).Length > 0);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

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

    private sealed class CancelWhenOutputBeginsStream(CancellationTokenSource cancellation) : MemoryStream {
        public bool OutputBegan { get; private set; }
        public override bool CanWrite {
            get {
                OutputBegan = true;
                cancellation.Cancel();
                return base.CanWrite;
            }
        }
    }
}
