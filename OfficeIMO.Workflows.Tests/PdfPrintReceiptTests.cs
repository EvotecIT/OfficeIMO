namespace OfficeIMO.Workflows.Tests;

public sealed class PdfPrintReceiptTests {
    [Fact]
    public void CleanupFailurePreservesAcceptedReceiptAndReportsTheStagingPath() {
        var receipt = new PdfPrintSubmission("queue", "queue-42", 2, 3, null);
        PdfPrintSubmission result = CupsPdfPrinter.CleanupStaging("private-staging", receipt, null, true, receipt.JobId,
            _ => throw new UnauthorizedAccessException("Access denied"))!;
        Assert.Equal(receipt.JobId, result.JobId);
        Assert.Equal(2, result.SheetCount);
        Assert.Equal(3, result.Copies);
        Assert.Contains("private-staging", result.CleanupWarning);
        Assert.Contains("Access denied", result.CleanupWarning);
    }

    [Fact]
    public void CleanupFailureRetainsUncertainSubmissionAndItsOriginalFailure() {
        var original = new IOException("No acknowledgement");
        PdfPrintDeliveryException error = Assert.Throws<PdfPrintDeliveryException>(() =>
            CupsPdfPrinter.CleanupStaging("private-staging", null, original, true, null, _ => throw new IOException("Locked")));
        var details = Assert.IsType<AggregateException>(error.InnerException);
        Assert.Contains(original, details.InnerExceptions);
        Assert.Null(error.JobId);
    }

    [Fact]
    public void SuccessfulCleanupLeavesReceiptUnchanged() {
        var receipt = new PdfPrintSubmission("queue", "queue-42", 2, 1, null);
        string? deleted = null;
        Assert.Same(receipt, CupsPdfPrinter.CleanupStaging("private-staging", receipt, null, true, receipt.JobId, path => deleted = path));
        Assert.Equal("private-staging", deleted);
    }
}
