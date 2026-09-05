using System;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfConversionCancellationTests {
    [Fact]
    public void OpenedPdfHtmlConversionAppliesMethodCancellationBeforeSemanticRead() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Cancelled HTML conversion must not start semantic reading"))
            .ToBytes();
        var glyphStage = new TrackingGlyphStage();
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        PdfToHtmlOptions options = CreateHtmlOptions(glyphStage);

        Assert.Throws<OperationCanceledException>(() => PdfDocument.Load(source).ToHtml(options, cancellation.Token));

        Assert.False(glyphStage.WasCalled);
    }

    [Fact]
    public void OpenedPdfHtmlResultAppliesMethodCancellationBeforeSemanticRead() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Cancelled HTML result must not start semantic reading"))
            .ToBytes();
        var glyphStage = new TrackingGlyphStage();
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        PdfToHtmlOptions options = CreateHtmlOptions(glyphStage);

        Assert.Throws<OperationCanceledException>(() => PdfDocument.Load(source).ToHtmlResult(options, cancellation.Token));

        Assert.False(glyphStage.WasCalled);
    }

    [Fact]
    public async Task OpenedPdfHtmlExportAppliesMethodCancellationBeforeSemanticRead() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Cancelled HTML export must not start semantic reading"))
            .ToBytes();
        var glyphStage = new TrackingGlyphStage();
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        PdfToHtmlOptions options = CreateHtmlOptions(glyphStage);

        await Assert.ThrowsAsync<OperationCanceledException>(() =>
            PdfDocument.Load(source).SaveAsHtmlAsync(Stream.Null, options, cancellation.Token));

        Assert.False(glyphStage.WasCalled);
    }

    private static PdfToHtmlOptions CreateHtmlOptions(
        IPdfGlyphDecodingStage glyphStage) => new() {
            ReadOptions = new PdfReadOptions {
                Profile = PdfReadProfile.Structured,
                Pipeline = new PdfUnderstandingPipelineOptions { GlyphDecoding = glyphStage }
            }
        };

    [Fact]
    public void SynchronousPdfImportsPassMethodCancellationIntoTheInitialSemanticRead() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Cancellation must reach the semantic read"))
            .ToBytes();
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() => {
            _ = PdfDocument.Load(source).ToWordDocumentResult(new PdfToWordOptions(), cancellation.Token);
        });
        Assert.Throws<OperationCanceledException>(() => {
            _ = PdfDocument.Load(source).ImportTablesToExcelDocumentResult(new PdfTablesToExcelOptions(), cancellation.Token);
        });
        Assert.Throws<OperationCanceledException>(() => {
            _ = PdfDocument.Load(source).ToPowerPointPresentationResult(new PdfToPowerPointOptions(), cancellation.Token);
        });
    }

    [Fact]
    public void InspectionAndTableContinuationOwnersExposeCancellation() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Cancellation must reach inspection and table grouping"))
            .ToBytes();
        PdfDocument document = PdfDocument.Load(source);
        PdfDocumentReadResult logical = document.Read();
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() => document.Inspect(null, cancellation.Token));
        Assert.Throws<OperationCanceledException>(() =>
            logical.GetTableContinuationGroups(options: null, cancellationToken: cancellation.Token));
    }

    private sealed class TrackingGlyphStage : IPdfGlyphDecodingStage {
        internal bool WasCalled { get; private set; }

        public IReadOnlyList<PdfTextSpan> Decode(PdfUnderstandingPageContext context) {
            WasCalled = true;
            return Array.Empty<PdfTextSpan>();
        }
    }
}
