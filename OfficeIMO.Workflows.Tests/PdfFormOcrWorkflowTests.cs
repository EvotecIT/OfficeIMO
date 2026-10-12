using OfficeIMO.Ocr;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public sealed class PdfFormOcrWorkflowTests {
    [Fact]
    public async Task PreparationEnforcesWorkflowInputAdmissionBeforeCallingProvider() {
        var document = PdfDocument.Create(compose => compose.Page(page => page.Size(300, 300)));
        byte[] source = document.ToBytes();
        int calls = 0;
        var engine = new DelegateOcrEngine("admission-fixture", (_, _) => { calls++; return Task.FromResult(new OcrResult()); });
        var runner = new OfficeWorkflowRunner();
        await Assert.ThrowsAsync<InvalidOperationException>(() => runner.PreparePdfFormOcrAsync(document, engine,
            limits: new OfficeWorkflowLimits { MaximumInputBytes = source.Length - 1 }));
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => runner.PreparePdfFormOcrAsync(document, engine,
            cancellationToken: new CancellationToken(true)));
        Assert.Equal(0, calls); Assert.Equal(source, document.ToBytes());
        var review = await runner.PreparePdfFormOcrAsync(document, engine,
            limits: new OfficeWorkflowLimits { MaximumInputBytes = source.Length });
        Assert.Equal(1, calls); Assert.Empty(review.Proposals); Assert.Equal(source, document.ToBytes());
    }
}
