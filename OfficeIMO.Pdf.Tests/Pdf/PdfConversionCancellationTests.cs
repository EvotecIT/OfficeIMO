using System;
using System.Threading;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfConversionCancellationTests {
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
}
