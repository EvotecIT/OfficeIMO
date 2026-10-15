using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Reader.Tests;

public sealed class ReaderAdmissionLimitPipelineTests {
    [Theory]
    [InlineData(OfficeDocumentProcessorFailureBehavior.Throw, false)]
    [InlineData(OfficeDocumentProcessorFailureBehavior.StopWithDiagnostic, false)]
    [InlineData(OfficeDocumentProcessorFailureBehavior.ContinueWithDiagnostic, false)]
    [InlineData(OfficeDocumentProcessorFailureBehavior.Throw, true)]
    [InlineData(OfficeDocumentProcessorFailureBehavior.StopWithDiagnostic, true)]
    [InlineData(OfficeDocumentProcessorFailureBehavior.ContinueWithDiagnostic, true)]
    public async Task NativeAdmissionCannotBecomeAProcessorDiagnosticOrWrapper(
        OfficeDocumentProcessorFailureBehavior behavior, bool asynchronous) {
        var native = new InvalidDataException("Native admission cause");
        var admission = new ReaderAdmissionLimitException("OpenDocument.MaxCompressionRatio", 1.5, 2.75, native);
        bool laterRan = false;
        var pipeline = new OfficeDocumentProcessorPipelineBuilder()
            .Add(new DelegateOfficeDocumentProcessor("nested-admission", (_, _) => throw admission))
            .Add(new DelegateOfficeDocumentProcessor("later", (document, _) => { laterRan = true; return document; }))
            .Build();
        var options = new OfficeDocumentProcessingOptions { FailureBehavior = behavior };
        ReaderAdmissionLimitException result = asynchronous
            ? await Assert.ThrowsAsync<ReaderAdmissionLimitException>(() => pipeline.ProcessAsync(new OfficeDocumentReadResult(), options))
            : Assert.Throws<ReaderAdmissionLimitException>(() => pipeline.Process(new OfficeDocumentReadResult(), options));
        Assert.Same(admission, result);
        Assert.Same(native, result.InnerException);
        Assert.Equal(1.5, result.Maximum);
        Assert.Equal(2.75, result.Actual);
        Assert.False(laterRan);
    }
}
