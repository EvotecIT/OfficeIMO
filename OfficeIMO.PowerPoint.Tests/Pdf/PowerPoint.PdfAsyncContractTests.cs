using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PowerPointPdfAsyncContractTests {
    [Fact]
    public async Task PdfAsyncSavesPerformIoAndDoNotTurnCancellationIntoFailureResults() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create();
        presentation.AddSlide().AddTextBoxPoints("Async PowerPoint PDF", 36, 36, 240, 48);
        using var output = new MemoryStream();

        await presentation.SaveAsPdfAsync(output);

        Assert.Equal("%PDF-", System.Text.Encoding.ASCII.GetString(output.ToArray(), 0, 5));
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            presentation.SaveAsPdfAsync(new MemoryStream(), cancellationToken: cancellation.Token));
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            presentation.SaveAsPdfResultAsync(new MemoryStream(), cancellationToken: cancellation.Token));
    }

    [Fact]
    public async Task PdfImportAsyncPassesCancellationIntoSemanticReconstruction() {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Cancellation probe"))
            .ToBytes();
        using var cancellation = new CancellationTokenSource();
        var stage = new CancelingSemanticStage(cancellation);
        var options = PdfToPowerPointOptions.CreateEditableTables();
        options.ReadOptions = new PdfReadOptions {
            Pipeline = new PdfUnderstandingPipelineOptions { SemanticClassification = stage }
        };
        using var output = new MemoryStream();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            PdfDocument.Load(pdf).SaveAsPowerPointAsync(output, options, cancellation.Token));

        Assert.True(stage.ObservedSemanticCancellation);
    }

    [Fact]
    public async Task PdfImportAsyncPassesMethodCancellationIntoSaveStage() {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Save cancellation probe"))
            .ToBytes();
        using var cancellation = new CancellationTokenSource();
        var options = PdfToPowerPointOptions.CreateEditableContent();
        using var output = new CancelOnWriteStream(cancellation);

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            PdfDocument.Load(pdf).SaveAsPowerPointAsync(output, options, cancellation.Token));

        Assert.True(output.ObservedOptionToken);
    }

    [Fact]
    public void LogicalPdfImportHonorsMethodCancellation() {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Logical cancellation probe"))
            .ToBytes();
        PdfDocumentReadResult logical = PdfDocument.Load(pdf).Read();
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        var options = PdfToPowerPointOptions.CreateEditableTables();

        Assert.ThrowsAny<OperationCanceledException>(() =>
            logical.ToPowerPointPresentationResult(options, cancellation.Token));
    }

    private sealed class CancelingSemanticStage : IPdfSemanticClassificationStage {
        private readonly CancellationTokenSource _cancellation;

        internal CancelingSemanticStage(CancellationTokenSource cancellation) {
            _cancellation = cancellation;
        }

        internal bool ObservedSemanticCancellation { get; private set; }

        public IReadOnlyList<PdfUnderstandingSemanticElement> Classify(
            PdfUnderstandingPageContext context,
            IReadOnlyList<PdfUnderstandingRegion> orderedRegions) {
            _cancellation.Cancel();
            try {
                context.ThrowIfCancellationRequested();
            } catch (OperationCanceledException) {
                ObservedSemanticCancellation = true;
                throw;
            }

            return Array.Empty<PdfUnderstandingSemanticElement>();
        }
    }

    private sealed class CancelOnWriteStream : MemoryStream {
        private readonly CancellationTokenSource _cancellation;

        internal CancelOnWriteStream(CancellationTokenSource cancellation) {
            _cancellation = cancellation;
        }

        internal bool ObservedOptionToken { get; private set; }

        public override Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken) {
            _cancellation.Cancel();
            ObservedOptionToken = cancellationToken == _cancellation.Token;
            cancellationToken.ThrowIfCancellationRequested();
            return base.WriteAsync(buffer, offset, count, cancellationToken);
        }

#if NET6_0_OR_GREATER
        public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken cancellationToken = default) {
            _cancellation.Cancel();
            ObservedOptionToken = cancellationToken == _cancellation.Token;
            cancellationToken.ThrowIfCancellationRequested();
            return base.WriteAsync(buffer, cancellationToken);
        }
#endif
    }
}
