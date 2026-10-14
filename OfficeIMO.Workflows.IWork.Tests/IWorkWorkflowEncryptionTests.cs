using OfficeIMO.Pdf;
using OfficeIMO.Workflows.IWork;
using Xunit;

namespace OfficeIMO.Workflows.IWork.Tests;

public sealed partial class IWorkWorkflowTests {
    [Theory]
    [InlineData("pages", "pages-pdf", false)]
    [InlineData("pages", "pages-pdf", true)]
    [InlineData("numbers", "numbers-pdf", false)]
    [InlineData("numbers", "numbers-pdf", true)]
    [InlineData("key", "keynote-pdf", false)]
    [InlineData("key", "keynote-pdf", true)]
    public async Task Registered_pdf_encryption_is_rejected_before_capture_or_publication(string kind, string route, bool compress) {
        using var files = new Files(kind, "pdf");
        File.WriteAllText(files.Output, "keep existing");
        var settings = new IWorkWorkflowSettings();
        var pdf = new PdfOptions().SetEncryption("synthetic-document-password");
        if (kind == "pages") settings.PagesPdfOptions.PdfOptions = pdf;
        else if (kind == "numbers") settings.NumbersPdfOptions.PdfOptions = pdf;
        else settings.KeynotePdfOptions.PdfOptions = pdf;
        OfficeWorkflowRequest request = files.Request(route);
        request.RegisteredConversionSettings = settings;
        request.ConversionOptions = new() { CompressPdfOutput = compress };
        request.ConflictPolicy = OfficeWorkflowConflictPolicy.Replace;
        var result = await IWorkWorkflow.CreateRunner().RunAsync(request);
        Assert.False(result.Succeeded);
        Assert.Contains("encryption", result.Summary, StringComparison.OrdinalIgnoreCase);
        Assert.Equal("validate", Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "WorkflowFailed").Stage);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "SourceSnapshot");
        Assert.Equal("keep existing", File.ReadAllText(files.Output));
        Assert.Empty(Directory.GetFiles(files.Root, "*.tmp"));
    }
}
