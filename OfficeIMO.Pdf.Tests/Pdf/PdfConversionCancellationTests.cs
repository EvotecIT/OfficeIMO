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
    public void OpenedPdfHtmlConversionAppliesOptionsCancellationBeforeSemanticRead() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Cancelled HTML conversion must not start semantic reading"))
            .ToBytes();
        var glyphStage = new TrackingGlyphStage();
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        PdfHtmlSaveOptions options = CreateCancelledHtmlOptions(cancellation.Token, glyphStage);

        Assert.Throws<OperationCanceledException>(() => PdfDocument.Load(source).ToHtml(options));

        Assert.False(glyphStage.WasCalled);
    }

    [Fact]
    public void OpenedPdfHtmlResultAppliesOptionsCancellationBeforeSemanticRead() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Cancelled HTML result must not start semantic reading"))
            .ToBytes();
        var glyphStage = new TrackingGlyphStage();
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        PdfHtmlSaveOptions options = CreateCancelledHtmlOptions(cancellation.Token, glyphStage);

        Assert.Throws<OperationCanceledException>(() => PdfDocument.Load(source).ToHtmlResult(options));

        Assert.False(glyphStage.WasCalled);
    }

    [Fact]
    public async Task OpenedPdfHtmlExportAppliesOptionsCancellationBeforeSemanticRead() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Cancelled HTML export must not start semantic reading"))
            .ToBytes();
        var glyphStage = new TrackingGlyphStage();
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        PdfHtmlSaveOptions options = CreateCancelledHtmlOptions(cancellation.Token, glyphStage);

        await Assert.ThrowsAsync<OperationCanceledException>(() =>
            PdfDocument.Load(source).SaveAsHtmlAsync(Stream.Null, options));

        Assert.False(glyphStage.WasCalled);
    }

    private static PdfHtmlSaveOptions CreateCancelledHtmlOptions(
        CancellationToken cancellationToken,
        IPdfGlyphDecodingStage glyphStage) => new() {
            CancellationToken = cancellationToken,
            ReadOptions = new PdfReadOptions {
                Profile = PdfReadProfile.Structured,
                Pipeline = new PdfUnderstandingPipelineOptions { GlyphDecoding = glyphStage }
            }
        };

    [Fact]
    public void SynchronousPdfImportsPassOptionCancellationIntoTheInitialSemanticRead() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Cancellation must reach the semantic read"))
            .ToBytes();
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() => {
            _ = PdfDocument.Load(source).ToWordDocumentResult(new PdfWordImportOptions {
                CancellationToken = cancellation.Token
            });
        });
        Assert.Throws<OperationCanceledException>(() => {
            _ = PdfDocument.Load(source).ImportTablesToExcelDocumentResult(new PdfExcelTableImportOptions {
                CancellationToken = cancellation.Token
            });
        });
        Assert.Throws<OperationCanceledException>(() => {
            _ = PdfDocument.Load(source).ToPowerPointPresentationResult(new PdfPowerPointImportOptions {
                CancellationToken = cancellation.Token
            });
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
