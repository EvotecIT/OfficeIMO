using OfficeIMO.Chm;
using OfficeIMO.Chm.Tests;
using OfficeIMO.Epub;
using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows.Tests;

public sealed class ChmWorkflowTests {
    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    public async Task AssemblyCombinesSingleAndMultiTopicHelpWithOtherDocuments(int topicCount) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-chm-assembly-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string input = Path.Combine(root, "source.chm"), text = Path.Combine(root, "other.pdf"), output = Path.Combine(root, "assembled.pdf");
            var topics = Enumerable.Range(1, topicCount).ToDictionary(index => "/topic" + index + ".html",
                index => ChmFixture.Html("<p>Help topic " + index + "</p>"));
            File.WriteAllBytes(input, ChmFixture.Archive(topics));
            File.WriteAllBytes(text, PdfDocument.Create(compose => compose.Page(page => page.Content(content =>
                content.Item(item => item.Paragraph(paragraph => paragraph.Text("Other document")))))).ToBytes());
            var result = await new OfficeWorkflowRunner().AssemblePdfAsync(new PdfAssemblyRequest { Sources = [input, text], OutputPath = output });
            Assert.True(result.Succeeded, result.Summary);
            var pdf = PdfReadDocument.Open(File.ReadAllBytes(output));
            Assert.Equal(topicCount + 1, pdf.Pages.Count);
            string content = pdf.ExtractText();
            for (int index = 1; index <= topicCount; index++) Assert.Contains("Help topic " + index, content);
            Assert.Contains("Other document", content);
            Assert.Empty(Directory.GetFiles(root, ".*.tmp"));
        } finally { Directory.Delete(root, true); }
    }

    [Theory]
    [InlineData(".md", "chm-markdown")]
    [InlineData(".epub", "chm-epub")]
    [InlineData(".pdf", "chm-pdf")]
    public async Task RegisteredRoutesPublishReopenedArtifactsAndRetainEvidence(string extension, string route) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-chm-workflow-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            byte[] source = ChmFixture.Book(); string input = Path.Combine(root, "source.chm"), output = Path.Combine(root, "result" + extension);
            File.WriteAllBytes(input, source);
            Assert.Equal(route, OfficeWorkflowCatalog.Find(".chm", extension)?.Id);
            OfficeWorkflowResult result = await OfficeWorkflow.Convert(input).To(output).RunAsync();
            Assert.True(result.Succeeded, result.Summary);
            var evidence = Assert.IsType<OfficeWorkflowConversionEvidence>(result.ConversionEvidence);
            Assert.Equal("ITSF", evidence.Facts["sourceContainer"]); Assert.Equal("2", evidence.Facts["topicCount"]);
            Assert.Contains(result.Diagnostics, item => item.Code == "OutputReopened");
            byte[] bytes = File.ReadAllBytes(output);
            if (extension == ".pdf") { Assert.Equal(2, PdfReadDocument.Open(bytes).Pages.Count); Assert.Contains("Café", PdfReadDocument.Open(bytes).ExtractText()); }
            else if (extension == ".epub") { using var stream = new MemoryStream(bytes); Assert.Contains(EpubDocument.Load(stream).Chapters, item => item.Text.Contains("Café")); }
            else Assert.Contains("Café", File.ReadAllText(output));
            Assert.Equal(source, File.ReadAllBytes(input)); Assert.Empty(Directory.GetFiles(root, ".*.tmp"));
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public async Task ExhaustedConversionBudgetLeavesExistingDestinationIntact() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-chm-workflow-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string input = Path.Combine(root, "source.chm"), output = Path.Combine(root, "result.md");
            byte[] sentinel = [3, 2, 1]; File.WriteAllBytes(input, ChmFixture.Book()); File.WriteAllBytes(output, sentinel);
            var result = await new OfficeWorkflowRunner().RunAsync(new OfficeWorkflowRequest {
                Operation = OfficeWorkflowOperation.Convert, InputPath = input, OutputPath = output,
                ConversionRouteId = "chm-markdown", ConflictPolicy = OfficeWorkflowConflictPolicy.Replace,
                ConversionOptions = new() { Chm = new() { MaxTotalHtmlCharacters = 10 } }
            });
            Assert.False(result.Succeeded); Assert.Equal(sentinel, File.ReadAllBytes(output));
            Assert.Empty(Directory.GetFiles(root, ".*.tmp"));
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public void TopicSelectionIsDetachedAndRouteSpecific() {
        var options = new OfficeWorkflowConversionOptions { Chm = new() { TopicPaths = new[] { "/welcome.html" } } };
        var snapshot = options.ForRoute("chm-epub"); options.Chm.TopicPaths = new[] { "/missing.html" };
        Assert.Equal("/welcome.html", Assert.Single(snapshot.Chm!.TopicPaths!)); Assert.Null(options.ForRoute("html-pdf").Chm);
    }
}
