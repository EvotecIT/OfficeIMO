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
        var request = ((new OfficeWorkflowRequest {
            Operation = OfficeWorkflowOperation.Convert,
            InputPath = input,
            OutputPath = output
,
            ConversionRouteId = OfficeWorkflowCatalog.Find(Path.GetExtension(input), Path.GetExtension(output), executableOnly: true)?.Id
        }));
        Assert.Equal("xps-pdf", OfficeWorkflowCatalog.Find(extension, ".pdf", executableOnly: true)?.Id);
        var result = await new OfficeWorkflowRunner().RunAsync(request);
        Assert.True(result.Status == OfficeWorkflowStatus.Completed, result.Summary);
        var pdf = PdfDocument.Load(output); Assert.Equal(1, pdf.Inspect().PageCount);
        Assert.Contains("Native workflow text", pdf.Read().Text);
        Assert.Contains(result.Diagnostics, d => d.Code == "conversion-output-reopened");
    }

    [Fact]
    public async Task OutputBudgetRejectsSerializationWithoutReplacingAnExistingFile() {
        using var scope = new Scope(); string input = scope.File("source.oxps"), output = scope.File("result.pdf");
        Create(XpsFormat.OpenXps).Save(input); File.WriteAllText(output, "Existing output");
        var result = await new OfficeWorkflowRunner().RunAsync(((new OfficeWorkflowRequest {
            Operation = OfficeWorkflowOperation.Convert,
            InputPath = input,
            OutputPath = output,
            ConflictPolicy = OfficePublicationPolicy.Replace,
            Limits = new OfficeOperationLimits {
                MaximumInputBytes = 1024 * 1024,
                MaximumOutputBytes = 64
            }
,
            ConversionRouteId = OfficeWorkflowCatalog.Find(Path.GetExtension(input), Path.GetExtension(output), executableOnly: true)?.Id
        })));
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
        Assert.Equal(OfficeWorkflowStatus.Failed, (await runner.RunAsync(((new OfficeWorkflowRequest {
            Operation = OfficeWorkflowOperation.Convert,
            InputPath = input,
            OutputPath = output
,
            ConversionRouteId = OfficeWorkflowCatalog.Find(Path.GetExtension(input), Path.GetExtension(output), executableOnly: true)?.Id
        })))).Status);
        Assert.False(File.Exists(output));
        var result = await runner.RunAsync(((new OfficeWorkflowRequest {
            Operation = OfficeWorkflowOperation.Convert,
            InputPath = input,
            OutputPath = output,
            ConversionOptions = (new OfficeConversionOptions() {
                RegisteredSettings = new OfficeIMO.Workflows.OfficeWorkflowRouteSettings {
                    Xps = new XpsToPdfOptions {
                        PreserveLogicalStructure = false
                    }
                }
            }

    ).Clone()
,
            ConversionRouteId = OfficeWorkflowCatalog.Find(Path.GetExtension(input), Path.GetExtension(output), executableOnly: true)?.Id
        })));
        Assert.True(result.Status == OfficeWorkflowStatus.Completed, result.Summary);
        Assert.Contains("Native workflow text", PdfDocument.Load(output).Read().Text);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task NativeSecurityAndCompressionReachTheFinalArtifact(bool compression) {
        using var scope = new Scope(); string input = scope.File("source.oxps"), output = scope.File("result.pdf");
        Create(XpsFormat.OpenXps).Save(input);
        var result = await new OfficeWorkflowRunner().RunAsync(((new OfficeWorkflowRequest {
            Operation = OfficeWorkflowOperation.Convert,
            InputPath = input,
            OutputPath = output,
            ConversionOptions = (new OfficeConversionOptions() {
                CompressPdfOutput = compression,
                RegisteredSettings = new OfficeIMO.Workflows.OfficeWorkflowRouteSettings {
                    Xps = new XpsToPdfOptions {
                        PdfOptions = new PdfCreateOptions().SetEncryption("open")
                    }
                }
            }

    ).Clone()
,
            ConversionRouteId = OfficeWorkflowCatalog.Find(Path.GetExtension(input), Path.GetExtension(output), executableOnly: true)?.Id
        })));
        Assert.True(result.Status == OfficeWorkflowStatus.Completed, result.Summary);
        var pdf = PdfDocument.Load(output, new PdfLoadOptions { Password = "open" });
        Assert.Contains("Native workflow text", pdf.Read().Text);
        if (compression) Assert.Contains(result.Diagnostics, d => d.Code == "pdf-output-compressed");
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
    [InlineData(OfficeConversionProfile.Lightweight)]
    [InlineData(OfficeConversionProfile.PrintReady)]
    [InlineData(OfficeConversionProfile.TextOnly)]
    public async Task NativeAssemblyRejectsUnsupportedOutputProfiles(OfficeConversionProfile profile) {
        using var scope = new Scope(); string source = scope.File("source.oxps"), output = scope.File("result.pdf");
        Create(XpsFormat.OpenXps).Save(source); File.WriteAllText(output, "Existing output");
        var result = await new OfficeWorkflowRunner().AssemblePdfAsync(new() {
            Sources = [source],
            OutputPath = output,
            ConflictPolicy = OfficePublicationPolicy.Replace,
            OutputProfile = profile
        });
        Assert.Equal(OfficeWorkflowStatus.Failed, result.Status);
        Assert.Equal("Existing output", File.ReadAllText(output));
        Assert.Contains("xps-pdf", result.Summary);
        Assert.Contains("Faithful", result.Summary);
    }

    [Fact]
    public async Task TaggedNativeAssemblyPreservesBothSourcesAndTheirStructure() {
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
            Sources = [source, source],
            OutputPath = output,
            ConflictPolicy = OfficePublicationPolicy.Replace
        });
        Assert.True(result.Status == OfficeWorkflowStatus.Completed, result.Summary);
        var pdf = PdfDocument.Load(output);
        Assert.Equal(2, pdf.Inspect().PageCount);
        Assert.True(pdf.Inspect().HasTaggedContent);
        Assert.Equal(2, pdf.Inspect().TaggedContent!.StructureElements.Count(element => element.StructureType == "P"));
        Assert.All(pdf.Read().Pages, page => Assert.Contains("Native workflow text", string.Join("", page.TextBlocks.Select(block => block.Text))));
    }

    [Fact]
    public async Task CheckpointsReuseTheArtifactAndRejectChangedNativeSettings() {
        using var scope = new Scope(); string input = scope.File("inputs"); Directory.CreateDirectory(input);
        Create(XpsFormat.OpenXps).Save(Path.Combine(input, "source.oxps"));
        var request = new OfficeBatchOptions { InputDirectory = input, OutputDirectory = scope.File("outputs"), ResumeFrom = Path.Combine(scope.File("state"), "plan.json") };
        var runner = new OfficeWorkflowRunner();
        var first = await OfficeBatchExecutor.RunAsync(new OfficeConversionExecutor(), request); Assert.Equal(1, first.Completed); Assert.Equal(0, first.Failed);
        Assert.Equal(1, (await OfficeBatchExecutor.RunAsync(new OfficeConversionExecutor(), request)).Reused);
        var changed = await OfficeBatchExecutor.RunAsync(new OfficeConversionExecutor(), request with {
            ConversionOptions = new() {
                RegisteredSettings = new OfficeWorkflowRouteSettings { Xps = new() { PreserveLogicalStructure = false } }
            }
        });
        Assert.Equal(1, changed.Failed);
    }

    [Fact]
    public async Task CheckpointsHashNativePdfEncryptionSettingsWithoutPersistingPasswords() {
        using var scope = new Scope(); string input = scope.File("inputs"); Directory.CreateDirectory(input);
        Create(XpsFormat.Xps).Save(Path.Combine(input, "source.xps"));
        const string password = "Native checkpoint secret";
        var request = new OfficeBatchOptions {
            InputDirectory = input,
            OutputDirectory = scope.File("outputs"),
            ResumeFrom = Path.Combine(scope.File("state"), "plan.json"),
            ConversionOptions = new() {
                RegisteredSettings = new OfficeWorkflowRouteSettings { Xps = new() { PdfOptions = new PdfCreateOptions().SetEncryption(password) } }
            }
        };
        var runner = new OfficeWorkflowRunner();
        Assert.Equal(1, (await OfficeBatchExecutor.RunAsync(new OfficeConversionExecutor(), request)).Completed);
        Assert.Equal(1, (await OfficeBatchExecutor.RunAsync(new OfficeConversionExecutor(), request)).Reused);
        foreach (string file in Directory.EnumerateFiles(request.CheckpointDirectory!, "*", SearchOption.AllDirectories))
            Assert.DoesNotContain(password, File.ReadAllText(file));
        var changed = await OfficeBatchExecutor.RunAsync(new OfficeConversionExecutor(), request with {
            ConversionOptions = new() {
                RegisteredSettings = new OfficeWorkflowRouteSettings { Xps = new() { PdfOptions = new PdfCreateOptions().SetEncryption("Changed native secret") } }
            }
        });
        Assert.Equal(1, changed.Failed);
    }

    [Fact]
    public void RouteSettingsAreSnapshottedAndRemainSpecificToXps() {
        var options = new OfficeConversionOptions {
            RegisteredSettings = new OfficeWorkflowRouteSettings {
                Xps = new() { PreserveLogicalStructure = false, PdfOptions = new PdfCreateOptions().SetEncryption("original") }
            }
        };
        OfficeConversionOptions snapshot = options.Clone();
        var native = Assert.IsType<OfficeWorkflowRouteSettings>(options.RegisteredSettings).Xps!;
        native.PreserveLogicalStructure = true; native.PdfOptions!.SetEncryption("changed");
        var copied = Assert.IsType<OfficeWorkflowRouteSettings>(snapshot.RegisteredSettings).Xps!;
        Assert.False(copied.PreserveLogicalStructure); Assert.Equal("original", copied.PdfOptions!.Encryption!.UserPassword);
        Assert.Null(options.ForRoute("txt-pdf").RegisteredSettings);
        Assert.Throws<ArgumentException>(() => OfficeWorkflowRenderingOptions.Capture(options).Snapshot(OfficeWorkflowCatalog.FindExecutable("txt-pdf")!));
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