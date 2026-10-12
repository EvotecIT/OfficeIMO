using System.Security.Cryptography;
using Xunit;
using OfficeIMO.Workflows.IWork;

namespace OfficeIMO.Workflows.IWork.Tests;

public sealed partial class IWorkWorkflowTests {
    [Theory]
    [InlineData("docx", "pages-docx")]
    [InlineData("pdf", "pages-pdf")]
    public async Task Completed_output_metadata_survives_later_cancellation_and_external_removal(string extension, string route) {
        using var files = new Files("pages", extension);
        using var cancellation = new CancellationTokenSource();
        string? expectedHash = null;
        long expectedLength = 0;
        OfficeWorkflowRequest request = files.Request(route);
        if (extension == "pdf") request.RegisteredConversionSettings = new IWorkWorkflowSettings();
        var result = await IWorkWorkflow.CreateRunner().RunAsync(request, new InlineProgress(update => {
            if (update.Stage != "complete") return;
            byte[] bytes = File.ReadAllBytes(files.Output);
            expectedHash = Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
            expectedLength = bytes.LongLength;
            File.Delete(files.Output);
            cancellation.Cancel();
        }), cancellation.Token);
        Assert.True(cancellation.IsCancellationRequested, result.Summary);
        Assert.True(result.Succeeded, result.Summary);
        Assert.Equal(expectedLength, result.OutputBytes);
        Assert.NotNull(expectedHash);
        Assert.Equal(expectedHash, result.OutputSha256);
        Assert.False(File.Exists(files.Output));
    }
}
