using System.Xml.Linq;
using OfficeIMO.Pdf;
using OfficeIMO.Workflows;
using OfficeIMO.Xps;

namespace OfficeIMO.Workflows.Tests;

public sealed class XpsWorkflowTests {
    [Theory]
    [InlineData(XpsFormat.Xps, ".xps")]
    [InlineData(XpsFormat.OpenXps, ".oxps")]
    public async Task BothDialectsUseTheNativeRouteAndPublishReopenedPdf(XpsFormat format, string extension) {
        using var scope = new Scope();
        string input = scope.File("source" + extension), output = scope.File("result.pdf");
        Create(format).Save(input);
        var request = OfficeWorkflow.Convert(input).To(output).Build();
        Assert.Equal("xps-pdf", request.ConversionRouteId);
        var result = await new OfficeWorkflowRunner().RunAsync(request);
        Assert.True(result.Status == OfficeWorkflowStatus.Completed, result.Summary);
        var pdf = PdfDocument.Load(output); Assert.Equal(1, pdf.Inspect().PageCount);
        Assert.Contains("Native workflow text", pdf.Read().Text);
        Assert.Contains(result.Diagnostics, d => d.Code == "OutputReopened");
    }

    [Fact]
    public async Task OutputBudgetRejectsSerializationWithoutReplacingAnExistingFile() {
        using var scope = new Scope(); string input = scope.File("source.oxps"), output = scope.File("result.pdf");
        Create(XpsFormat.OpenXps).Save(input); File.WriteAllText(output, "Existing output");
        var result = await new OfficeWorkflowRunner().RunAsync(OfficeWorkflow.Convert(input).To(output)
            .OnConflict(OfficeWorkflowConflictPolicy.Replace).WithLimits(1024 * 1024, 64).Build());
        Assert.Equal(OfficeWorkflowStatus.Failed, result.Status);
        Assert.Equal("Existing output", File.ReadAllText(output));
        Assert.Contains(result.Diagnostics, d => d.Message.Contains("while it was being serialized", StringComparison.Ordinal));
    }

    [Fact]
    public async Task UnsupportedNativeSemanticsRequireTheExplicitOptOut() {
        using var scope = new Scope(); var document = Create(XpsFormat.Xps);
        XNamespace ns = "http://schemas.microsoft.com/xps/2005/06/documentstructure";
        document.Pages[0].ReplaceStoryFragmentsMarkup(new XElement(ns + "StoryFragments", new XElement(ns + "StoryFragment",
            new XAttribute("FragmentType", "Content"), new XElement(XName.Get("Semantic", "urn:extension")))));
        string input = scope.File("source.xps"), output = scope.File("result.pdf"); document.Save(input);
        var runner = new OfficeWorkflowRunner();
        Assert.Equal(OfficeWorkflowStatus.Failed, (await runner.RunAsync(OfficeWorkflow.Convert(input).To(output).Build())).Status);
        Assert.False(File.Exists(output));
        var result = await runner.RunAsync(OfficeWorkflow.Convert(input).To(output).WithConversionOptions(new() {
            Xps = new XpsToPdfOptions { PreserveLogicalStructure = false }
        }).Build());
        Assert.True(result.Status == OfficeWorkflowStatus.Completed, result.Summary);
        Assert.Contains("Native workflow text", PdfDocument.Load(output).Read().Text);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task NativeSecurityAndCompressionReachTheFinalArtifact(bool compression) {
        using var scope = new Scope(); string input = scope.File("source.oxps"), output = scope.File("result.pdf");
        Create(XpsFormat.OpenXps).Save(input);
        var result = await new OfficeWorkflowRunner().RunAsync(OfficeWorkflow.Convert(input).To(output).WithConversionOptions(new() {
            CompressPdfOutput = compression,
            Xps = new XpsToPdfOptions { PdfOptions = new PdfOptions().SetEncryption("open") }
        }).Build());
        Assert.True(result.Status == OfficeWorkflowStatus.Completed, result.Summary);
        var pdf = PdfDocument.Load(output, new PdfLoadOptions { Password = "open" });
        Assert.Contains("Native workflow text", pdf.Read().Text);
        if (compression) Assert.Contains(result.Diagnostics, d => d.Code == "PdfOutputCompression");
    }

    [Fact]
    public async Task PdfAssemblyAcceptsBothNativeDialectSources() {
        using var scope = new Scope(); string legacy = scope.File("one.xps"), open = scope.File("two.oxps");
        Create(XpsFormat.Xps).Save(legacy); Create(XpsFormat.OpenXps).Save(open);
        var result = await new OfficeWorkflowRunner().AssemblePdfAsync(new() { Sources = [legacy, open], OutputPath = scope.File("combined.pdf") });
        Assert.True(result.Status == OfficeWorkflowStatus.Completed, result.Summary); Assert.Equal(2, result.PageCount);
        Assert.Equal(2, PdfDocument.Load(result.OutputPath!).Inspect().PageCount);
    }

    [Theory]
    [InlineData(OfficeWorkflowOutputProfile.Lightweight)]
    [InlineData(OfficeWorkflowOutputProfile.PrintReady)]
    [InlineData(OfficeWorkflowOutputProfile.TextOnly)]
    public async Task NativeAssemblyRejectsUnsupportedOutputProfiles(OfficeWorkflowOutputProfile profile) {
        using var scope = new Scope(); string source = scope.File("source.oxps"), output = scope.File("result.pdf");
        Create(XpsFormat.OpenXps).Save(source); File.WriteAllText(output, "Existing output");
        var result = await new OfficeWorkflowRunner().AssemblePdfAsync(new() {
            Sources = [source], OutputPath = output, ConflictPolicy = OfficeWorkflowConflictPolicy.Replace, OutputProfile = profile
        });
        Assert.Equal(OfficeWorkflowStatus.Failed, result.Status);
        Assert.Equal("Existing output", File.ReadAllText(output));
        Assert.Contains("xps-pdf", result.Summary);
        Assert.Contains("Faithful", result.Summary);
    }

    [Fact]
    public async Task TaggedNativeAssemblyRejectsTheUnsupportedRewriteAndPreservesTheDestination() {
        using var scope = new Scope(); var document = Create(XpsFormat.OpenXps);
        var markup = document.Pages[0].GetMarkup(); markup.Elements().Single().SetAttributeValue("Name", "text");
        document.Pages[0].ReplaceMarkup(markup);
        XNamespace ns = "http://schemas.openxps.org/oxps/v1.0/documentstructure";
        document.Pages[0].ReplaceStoryFragmentsMarkup(new XElement(ns + "StoryFragments", new XElement(ns + "StoryFragment",
            new XAttribute("FragmentType", "Content"), new XElement(ns + "ParagraphStructure",
                new XElement(ns + "NamedElement", new XAttribute("NameReference", "text"))))));
        string source = scope.File("tagged.oxps"), output = scope.File("combined.pdf");
        document.Save(source); File.WriteAllText(output, "Existing output");
        var result = await new OfficeWorkflowRunner().AssemblePdfAsync(new() {
            Sources = [source, source], OutputPath = output, ConflictPolicy = OfficeWorkflowConflictPolicy.Replace
        });
        Assert.Equal(OfficeWorkflowStatus.Failed, result.Status);
        Assert.Contains("FullRewrite.TaggedContent", result.Summary);
        Assert.Equal("Existing output", File.ReadAllText(output));
    }

    [Fact]
    public async Task CheckpointsReuseTheArtifactAndRejectChangedNativeSettings() {
        using var scope = new Scope(); string input = scope.File("inputs"); Directory.CreateDirectory(input);
        Create(XpsFormat.OpenXps).Save(Path.Combine(input, "source.oxps"));
        var request = new OfficeConversionBatchRequest { InputDirectory = input, OutputDirectory = scope.File("outputs"), CheckpointDirectory = scope.File("state") };
        var runner = new OfficeWorkflowRunner();
        var first = await runner.RunBatchAsync(request); Assert.Equal(1, first.Completed); Assert.Equal(0, first.Failed);
        Assert.Equal(1, (await runner.RunBatchAsync(request)).Reused);
        var changed = await runner.RunBatchAsync(request with { ConversionOptions = new() { Xps = new() { PreserveLogicalStructure = false } } });
        Assert.Equal(1, changed.Failed);
    }

    [Fact]
    public async Task CheckpointsHashNativePdfEncryptionSettingsWithoutPersistingPasswords() {
        using var scope = new Scope(); string input = scope.File("inputs"); Directory.CreateDirectory(input);
        Create(XpsFormat.Xps).Save(Path.Combine(input, "source.xps"));
        const string password = "Native checkpoint secret";
        var request = new OfficeConversionBatchRequest {
            InputDirectory = input, OutputDirectory = scope.File("outputs"), CheckpointDirectory = scope.File("state"),
            ConversionOptions = new() { Xps = new() { PdfOptions = new PdfOptions().SetEncryption(password) } }
        };
        var runner = new OfficeWorkflowRunner();
        Assert.Equal(1, (await runner.RunBatchAsync(request)).Completed);
        Assert.Equal(1, (await runner.RunBatchAsync(request)).Reused);
        foreach (string file in Directory.EnumerateFiles(request.CheckpointDirectory, "*", SearchOption.AllDirectories))
            Assert.DoesNotContain(password, File.ReadAllText(file));
        var changed = await runner.RunBatchAsync(request with {
            ConversionOptions = new() { Xps = new() { PdfOptions = new PdfOptions().SetEncryption("Changed native secret") } }
        });
        Assert.Equal(1, changed.Failed);
    }

    [Fact]
    public void RouteSettingsAreSnapshottedAndRemainSpecificToXps() {
        var options = new OfficeWorkflowConversionOptions { Xps = new() { PreserveLogicalStructure = false, PdfOptions = new PdfOptions().SetEncryption("original") } };
        var builder = OfficeWorkflow.Convert("input.xps").To("result.pdf").WithConversionOptions(options);
        options.Xps.PreserveLogicalStructure = true; options.Xps.PdfOptions.SetEncryption("changed");
        var snapshot = builder.Build().ConversionOptions!.Xps!;
        Assert.False(snapshot.PreserveLogicalStructure); Assert.Equal("original", snapshot.PdfOptions!.Encryption!.UserPassword);
        Assert.Null(options.ForRoute("txt-pdf").Xps);
        Assert.Throws<ArgumentException>(() => options.Snapshot(OfficeWorkflowCatalog.FindExecutable("txt-pdf")!));
    }

    private static XpsDocument Create(XpsFormat format) {
        var document = XpsDocument.Create(format);
        string font = document.AddFont(File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "RobotoFlex.ttf")), false);
        document.AddPage(240, 180).AddText("Native workflow text", font, 14, 10, 25);
        return document;
    }

    private sealed class Scope : IDisposable {
        private readonly string _root = Path.Combine(Path.GetTempPath(), "officeimo-xps-workflow-" + Guid.NewGuid().ToString("N"));
        internal Scope() => Directory.CreateDirectory(_root);
        internal string File(string name) => Path.Combine(_root, name);
        public void Dispose() => Directory.Delete(_root, true);
    }
}
