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
            presentation.TrySaveAsPdfAsync(new MemoryStream(), cancellationToken: cancellation.Token));
    }

    [Fact]
    public async Task PdfImportAsyncPassesCancellationIntoSemanticReconstruction() {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Cancellation probe"))
            .ToBytes();
        using var cancellation = new CancellationTokenSource();
        var stage = new CancelingSemanticStage(cancellation);
        var options = PdfPowerPointImportOptions.CreateEditableTables();
        options.ReadOptions = new PdfReadOptions {
            Pipeline = new PdfUnderstandingPipelineOptions { SemanticClassification = stage }
        };
        using var output = new MemoryStream();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            PdfDocument.Load(pdf).SaveAsPowerPointAsync(output, options, cancellation.Token));

        Assert.True(stage.ObservedSemanticCancellation);
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
}
