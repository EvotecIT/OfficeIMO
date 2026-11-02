using System.Collections.Concurrent;
using OfficeIMO.Excel;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;
using ReadPdf = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public sealed class OfficeConversionBatchTests {
    [Fact]
    public async Task MixedFormatsUseTheCatalogAndReturnSkippedFilesAndVerifiedReuse() {
        using var scope = new BatchDirectory();
        File.WriteAllText(Path.Combine(scope.Input, "text.txt"), "Literal <h1> text");
        File.WriteAllText(Path.Combine(scope.Input, "markup.md"), "# Markdown source", System.Text.Encoding.Unicode);
        File.WriteAllText(Path.Combine(scope.Input, "page.html"), "<html><body><h1>HTML source</h1></body></html>");
        File.WriteAllText(Path.Combine(scope.Input, "rich.rtf"), "{\\rtf1\\ansi RTF source}");
        File.WriteAllText(Path.Combine(scope.Input, "other.bin"), "Unsupported input");
        using (var word = WordDocument.Create(Path.Combine(scope.Input, "word.docx"))) { word.AddParagraph("Word source"); word.Save(); }
        using (var excel = ExcelDocument.Create(Path.Combine(scope.Input, "sheet.xlsx"))) { excel.AddSheet("Data").Cell(1, 1, "Excel source"); excel.Save(); }
        using (var slides = PowerPointDocument.Create(Path.Combine(scope.Input, "slides.pptx"))) { slides.AddSlide().AddTextBox("Slide source",global::OfficeIMO.OfficeLength.Points(40),global::OfficeIMO.OfficeLength.Points(40),global::OfficeIMO.OfficeLength.Points(500),global::OfficeIMO.OfficeLength.Points(60)); slides.Save(); }
        var items = new Items();
        var builder = (((new OfficeBatchOptions { InputDirectory = scope.Input } with { OutputDirectory = scope.Output }) with { ResumeFrom = Path.Combine(scope.State, "plan.json"), RetryFailed = false }) with {
            ConversionOptions = new OfficeConversionOptions {
                Word = new OfficeIMO.Word.Pdf.WordPdfExportOptions { Title = "Batch Word" },
                Excel = new OfficeIMO.Excel.Pdf.ExcelPdfExportOptions { PdfOptions = new PdfCreateOptions { DefaultFontSize = global::OfficeIMO.OfficeLength.Points(12) } },
                PowerPoint = new OfficeIMO.PowerPoint.Pdf.PowerPointPdfExportOptions { PdfOptions = new PdfCreateOptions { DefaultFontSize = global::OfficeIMO.OfficeLength.Points(12) } },
                Html = new OfficeIMO.Html.Pdf.HtmlToPdfOptions(),
                Markdown = new OfficeIMO.Markdown.Pdf.MarkdownPdfExportOptions { Title = "Batch Markdown" },
                RegisteredSettings = new OfficeWorkflowRouteSettings { Rtf = new OfficeIMO.Rtf.Pdf.RtfToPdfOptions() }
            }.Clone()
        });
        var first = await OfficeBatchExecutor.RunAsync(new OfficeConversionExecutor(), builder, progress: items);
        Assert.Equal(7, first.Completed); Assert.Equal(0, first.Failed); Assert.Equal(1, first.Skipped);
        Assert.Single(items.Values, item => item.Skipped);
        foreach (var item in items.Values.Where(item => !item.Skipped)) {
            Assert.Equal(OfficeBatchItemStatus.Completed, item.Status);
            using var pdf = ReadPdf.Open(item.OutputPath!);
            Assert.NotEmpty(pdf.GetPages());
        }
        var second = await OfficeBatchExecutor.RunAsync(new OfficeConversionExecutor(), builder);
        Assert.Equal(7, second.Reused); Assert.Equal(1, second.Skipped);
    }

    [Fact]
    public async Task EncryptedOfficeFilesUseTheExistingLoadersAndPasswordsAreNotCheckpointed() {
        using var scope = new BatchDirectory();
        const string password = "Batch qualification secret";
        string wordPath = Path.Combine(scope.Input, "locked-word.docx");
        string excelPath = Path.Combine(scope.Input, "locked-excel.xlsx");
        using (var word = WordDocument.Create(wordPath)) { word.AddParagraph("Encrypted Word source"); word.Save(wordPath, new WordSaveOptions { Password = password }); }
        using (var excel = ExcelDocument.Create(excelPath)) { excel.AddSheet("Data").Cell(1, 1, "Encrypted Excel source"); excel.Save(excelPath, new ExcelSaveOptions { Password = password }); }
        var result = await OfficeBatchExecutor.RunAsync(new OfficeConversionExecutor(), (((new OfficeBatchOptions { InputDirectory = scope.Input } with { OutputDirectory = scope.Output }) with { ResumeFrom = Path.Combine(scope.State, "plan.json"), RetryFailed = false }) with {
            ConversionOptions = new OfficeConversionOptions {
                SourcePassword = password
            }.Clone()
        }));
        Assert.Equal(2, result.Completed); Assert.Equal(0, result.Failed);
        foreach (string checkpoint in Directory.EnumerateFiles(scope.State, "*.json", SearchOption.AllDirectories))
            Assert.DoesNotContain(password, File.ReadAllText(checkpoint));
    }

    [Fact]
    public async Task BatchTargetSelectionReusesTheExistingPdfToHtmlRoute() {
        using var scope = new BatchDirectory();
        string source = Path.Combine(scope.Input, "source.pdf");
        PdfPlainTextConverter.ToPdfDocumentResult("PDF batch source").Save(source);
        var result = await OfficeBatchExecutor.RunAsync(new OfficeConversionExecutor(), ((new OfficeBatchOptions { InputPaths = [source] } with { OutputDirectory = scope.Output, TargetExtension = ".html" }) with { ResumeFrom = Path.Combine(scope.State, "plan.json"), RetryFailed = false }));
        Assert.Equal(1, result.Completed);
        Assert.Contains("PDF batch source", File.ReadAllText(Path.Combine(scope.Output, "source.html")));
    }

    [Fact]
    public async Task ExistingTextOptionsApplyToSingleAndBatchAndExecutionBudgetsDoNotInvalidateReuse() {
        using var scope = new BatchDirectory();
        string source = Path.Combine(scope.Input, "geometry.txt");
        File.WriteAllText(source, "literal geometry");
        var text = new PdfTextImportOptions {
            PdfOptions = new PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(240),
                PageHeight = global::OfficeIMO.OfficeLength.Points(180),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(12),
                MarginRight = global::OfficeIMO.OfficeLength.Points(12),
                MarginTop = global::OfficeIMO.OfficeLength.Points(12),
                MarginBottom = global::OfficeIMO.OfficeLength.Points(12),
                DefaultFont = PdfStandardFont.Courier,
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(14)
            }
        };
        var options = new OfficeConversionOptions {
            PlainText = text
        };
        var single = await OfficeWorkflow.RunAsync((new OfficeWorkflowRequest {
            Operation = OfficeWorkflowOperation.Convert,
            InputPath = source,
            OutputPath = Path.Combine(scope.Output, "single.pdf"),
            ConversionOptions = (options).Clone()
,
            ConversionRouteId = OfficeWorkflowCatalog.Find(Path.GetExtension(source), Path.GetExtension(Path.Combine(scope.Output, "single.pdf")), executableOnly: true)?.Id
        }));
        Assert.True(single.Succeeded);
        var request = new OfficeBatchOptions {
            InputPaths = [source],
            OutputDirectory = scope.Output,
            ResumeFrom = Path.Combine(scope.State, "plan.json"),
            ConversionOptions = options,
            MaximumFiles = 1
        };
        Assert.Equal(1, (await OfficeBatchExecutor.RunAsync(new OfficeConversionExecutor(), request)).Completed);
        using (var pdf = ReadPdf.Open(Path.Combine(scope.Output, "geometry.pdf"))) {
            Assert.Equal(240, pdf.GetPage(1).Width); Assert.Equal(180, pdf.GetPage(1).Height);
        }
        var resumed = await OfficeBatchExecutor.RunAsync(new OfficeConversionExecutor(), request with {
            MaxParallelism = 9,
            MaximumFiles = 10,
            MaximumInputBytes = 128 * 1024 * 1024,
            MaximumOutputBytes = 512 * 1024 * 1024
        });
        Assert.Equal(1, resumed.Reused);
        text.PdfOptions.DefaultFontSize = global::OfficeIMO.OfficeLength.Points(18);
        Assert.Equal(1, (await OfficeBatchExecutor.RunAsync(new OfficeConversionExecutor(), request)).Failed);
    }

    [Fact]
    public async Task OrdinaryBatchesUseTheExistingRenameAndReplacePolicies() {
        using var scope = new BatchDirectory();
        string source = Path.Combine(scope.Input, "one.txt");
        File.WriteAllText(source, "First text");
        var builder = (new OfficeBatchOptions { InputPaths = [source] } with { OutputDirectory = scope.Output });
        Assert.Equal(1, (await OfficeBatchExecutor.RunAsync(new OfficeConversionExecutor(), builder)).Completed);
        Assert.Equal(1, (await OfficeBatchExecutor.RunAsync(new OfficeConversionExecutor(), (builder with { ConflictPolicy = OfficePublicationPolicy.Rename }))).Completed);
        Assert.Equal(2, Directory.GetFiles(scope.Output, "*.pdf").Length);
        File.WriteAllText(source, "Replacement text");
        Assert.Equal(1, (await OfficeBatchExecutor.RunAsync(new OfficeConversionExecutor(), (builder with { ConflictPolicy = OfficePublicationPolicy.Replace }))).Completed);
        using var pdf = ReadPdf.Open(Path.Combine(scope.Output, "one.pdf"));
        Assert.Contains("Replacement text", pdf.GetPage(1).Text);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task WordCheckpointsDistinguishImplicitWhitespaceFromExplicitCollapse(bool legacyBoolean) {
        using var scope = new BatchDirectory();
        string source = Path.Combine(scope.Input, "spaces.docx");
        using (var word = WordDocument.Create(source)) { word.AddParagraph("ALPHA   BETA"); word.Save(); }
        var pdfOptions = new PdfCreateOptions();
        var request = new OfficeBatchOptions {
            InputPaths = [source],
            OutputDirectory = scope.Output,
            ResumeFrom = Path.Combine(scope.State, "plan.json"),
            ConversionOptions = new() {
                Word = new() { IncludePageNumbers = false, PdfOptions = pdfOptions }
            }
        };
        var runner = new OfficeWorkflowRunner();
        Assert.Equal(1, (await OfficeBatchExecutor.RunAsync(new OfficeConversionExecutor(), request)).Completed);
        Assert.Equal(1, (await OfficeBatchExecutor.RunAsync(new OfficeConversionExecutor(), request)).Reused);
        byte[] original = File.ReadAllBytes(Path.Combine(scope.Output, "spaces.pdf"));
        if (legacyBoolean) pdfOptions.PreserveTextWhitespace = false;
        else pdfOptions.TextWhitespaceMode = PdfTextWhitespaceMode.Collapse;
        var changed = await OfficeBatchExecutor.RunAsync(new OfficeConversionExecutor(), request);
        Assert.Equal(0, changed.Reused);
        Assert.Equal(1, changed.Failed);
        Assert.Equal(original, File.ReadAllBytes(Path.Combine(scope.Output, "spaces.pdf")));
    }

    [Fact]
    public async Task HtmlCheckpointsIncludeLocalStylesheetContent() {
        using var scope = new BatchDirectory();
        File.WriteAllText(Path.Combine(scope.Input, "page.html"), "<html><head><link rel='stylesheet' href='style.css'></head><body><p>Styled source</p></body></html>");
        string css = Path.Combine(scope.Input, "style.css");
        File.WriteAllText(css, "p { font-size: 18px; }");
        var builder = ((new OfficeBatchOptions { InputDirectory = scope.Input } with { OutputDirectory = scope.Output }) with { ResumeFrom = Path.Combine(scope.State, "plan.json"), RetryFailed = false });
        Assert.Equal(1, (await OfficeBatchExecutor.RunAsync(new OfficeConversionExecutor(), builder)).Completed);
        Assert.Equal(1, (await OfficeBatchExecutor.RunAsync(new OfficeConversionExecutor(), builder)).Reused);
        byte[] original = File.ReadAllBytes(Path.Combine(scope.Output, "page.pdf"));
        File.WriteAllText(css, "p { font-size: 28px; }");
        Assert.Equal(1, (await OfficeBatchExecutor.RunAsync(new OfficeConversionExecutor(), builder)).Failed);
        Assert.Equal(original, File.ReadAllBytes(Path.Combine(scope.Output, "page.pdf")));
    }

    [Fact]
    public async Task DirectoryFilteringIsExplicitAndLiteralTextRemainsTheDefault() {
        using var scope = new BatchDirectory();
        File.WriteAllText(Path.Combine(scope.Input, "one.txt"), "<h1>literal text</h1>");
        File.WriteAllText(Path.Combine(scope.Input, "two.md"), "# Unselected");
        Directory.CreateDirectory(Path.Combine(scope.Input, "nested"));
        File.WriteAllText(Path.Combine(scope.Input, "nested", "three.txt"), "Nested text");
        var result = await OfficeBatchExecutor.RunAsync(new OfficeConversionExecutor(), ((new OfficeBatchOptions { InputDirectory = scope.Input } with { OutputDirectory = scope.Output }) with { Recursive = false, Extensions = [".txt"] }));
        Assert.Equal(1, result.Completed); Assert.Equal(1, result.Skipped);
        using var pdf = ReadPdf.Open(Path.Combine(scope.Output, "one.pdf"));
        Assert.Contains("<h1>literal text</h1>", pdf.GetPage(1).Text);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public async Task NativeOutputEncryptionWorksForEveryPdfBagWithOptionalCompressionAndCheckpoints(bool checkpoint, bool compress) {
        using var scope = new BatchDirectory();
        File.WriteAllText(Path.Combine(scope.Input, "text.txt"), "Encrypted text");
        File.WriteAllText(Path.Combine(scope.Input, "page.html"), "<p>Encrypted HTML</p>");
        File.WriteAllText(Path.Combine(scope.Input, "markup.md"), "# Encrypted Markdown");
        File.WriteAllText(Path.Combine(scope.Input, "rich.rtf"), "{\\rtf1\\ansi Encrypted RTF}");
        using (var word = WordDocument.Create(Path.Combine(scope.Input, "word.docx"))) { word.AddParagraph("Encrypted Word"); word.Save(); }
        using (var excel = ExcelDocument.Create(Path.Combine(scope.Input, "sheet.xlsx"))) { excel.AddSheet("Data").Cell(1, 1, "Encrypted Excel"); excel.Save(); }
        using (var slides = PowerPointDocument.Create(Path.Combine(scope.Input, "slides.pptx"))) { slides.AddSlide().AddTextBox("Encrypted slides",global::OfficeIMO.OfficeLength.Points(40),global::OfficeIMO.OfficeLength.Points(40),global::OfficeIMO.OfficeLength.Points(500),global::OfficeIMO.OfficeLength.Points(60)); slides.Save(); }
        const string userPassword = "Synthetic output reader", ownerPassword = "Synthetic output owner";
        PdfCreateOptions Encryption() => new PdfCreateOptions().SetEncryption(userPassword, ownerPassword);
        var request = new OfficeBatchOptions {
            InputDirectory = scope.Input,
            OutputDirectory = scope.Output,
            ResumeFrom = checkpoint ? Path.Combine(scope.State, "plan.json") : null,
            ConversionOptions = new() {
                PdfLoad = new() { Password = "Unrelated source password" },
                PlainText = new() { PdfOptions = Encryption() },
                Word = new() { PdfOptions = Encryption() },
                Excel = new() { PdfOptions = Encryption() },
                PowerPoint = new() { PdfOptions = Encryption() },
                Html = new() { PdfOptions = Encryption() },
                Markdown = new() { PdfOptions = Encryption() },
                CompressPdfOutput = compress,
                RegisteredSettings = new OfficeWorkflowRouteSettings { Rtf = new() { PdfOptions = Encryption() } }
            }
        };
        var items = new Items();
        var result = await OfficeBatchExecutor.RunAsync(new OfficeConversionExecutor(), request, items);
        Assert.True(result.Failed == 0, string.Join("\n", items.Values.Select(item => item.Summary))); Assert.Equal(7, result.Completed);
        foreach (var item in items.Values) {
            var pdf = PdfDocument.Load(item.OutputPath!, new PdfLoadOptions { Password = userPassword });
            Assert.True(pdf.Inspect().Security.HasEncryption); Assert.NotEmpty(pdf.Read().Pages);
        }
        if (checkpoint) {
            Assert.Equal(7, (await OfficeBatchExecutor.RunAsync(new OfficeConversionExecutor(), request)).Reused);
            foreach (string receipt in Directory.EnumerateFiles(scope.State, "*.json", SearchOption.AllDirectories)) {
                Assert.DoesNotContain(userPassword, File.ReadAllText(receipt));
                Assert.DoesNotContain(ownerPassword, File.ReadAllText(receipt));
            }
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task MarkdownImageBytesWithoutKnownExtensionsInvalidateReuseAndPublication(bool mutateAtPublication) {
        using var scope = new BatchDirectory();
        string input = Path.Combine(scope.Input, "page.md"), image = Path.Combine(scope.Input, "image.asset");
        File.WriteAllText(input, "# Image source\n![image](image.asset)");
        byte[] png = OcrSessionWorkflowTests.Png();
        File.WriteAllBytes(image, png);
        var request = new OfficeBatchOptions {
            InputPaths = [input],
            InputDirectory = scope.Input,
            OutputDirectory = scope.Output,
            ResumeFrom = Path.Combine(scope.State, "plan.json"),
            ConversionOptions = new() {
                Markdown = new() {
                    BaseDirectory = scope.Input,
                    ResourcePolicy = new PdfResourcePolicy { AllowLocalFileAccess = true }
                }
            }
        };
        if (mutateAtPublication) {
            var result = await OfficeBatchExecutor.RunAsync(new OfficeConversionExecutor(), request, publicationGuard: new ChangeResource(image));
            Assert.Equal(1, result.Failed); Assert.False(File.Exists(Path.Combine(scope.Output, "page.pdf")));
        } else {
            Assert.Equal(1, (await OfficeBatchExecutor.RunAsync(new OfficeConversionExecutor(), request)).Completed);
            string output = Path.Combine(scope.Output, "page.pdf");
            using (var pdf = ReadPdf.Open(output)) Assert.NotEmpty(pdf.GetPage(1).GetImages());
            byte[] original = File.ReadAllBytes(output);
            File.WriteAllBytes(image, [.. png, 0]);
            Assert.Equal(1, (await OfficeBatchExecutor.RunAsync(new OfficeConversionExecutor(), request)).Failed);
            Assert.Equal(original, File.ReadAllBytes(output));
        }
    }

    [Fact]
    public async Task UnrestrictedMarkdownResourcesRequireAnOrdinaryBatch() {
        using var scope = new BatchDirectory();
        string input = Path.Combine(scope.Input, "page.md");
        File.WriteAllText(input, "# Image source\n![image](../external.png)");
        File.WriteAllBytes(Path.Combine(Path.GetDirectoryName(scope.Input)!, "external.png"), OcrSessionWorkflowTests.Png());
        var request = new OfficeBatchOptions {
            InputPaths = [input],
            InputDirectory = scope.Input,
            OutputDirectory = scope.Output,
            ConversionOptions = new() {
                Markdown = new() {
                    BaseDirectory = scope.Input,
                    RestrictLocalImagesToBaseDirectory = false,
                    ResourcePolicy = new PdfResourcePolicy { AllowLocalFileAccess = true }
                }
            }
        };
        Assert.Equal(1, (await OfficeBatchExecutor.RunAsync(new OfficeConversionExecutor(), request)).Completed);
        using (var pdf = ReadPdf.Open(Path.Combine(scope.Output, "page.pdf"))) Assert.NotEmpty(pdf.GetPage(1).GetImages());
        var durable = await OfficeBatchExecutor.RunAsync(new OfficeConversionExecutor(), request with { ResumeFrom = Path.Combine(scope.State, "plan.json") });
        Assert.Equal(1, durable.Failed); Assert.Equal(0, durable.Reused);
    }

    [Theory]
    [InlineData(".md")]
    [InlineData(".html")]
    public async Task CorrectedFailedSourceCanRetryWithoutInvalidatingItsResourceConfiguration(string extension) {
        using var scope = new BatchDirectory();
        string input = Path.Combine(scope.Input, "page" + extension);
        // Distinct paragraphs keep the compressed PDF above the bound; repeated text compresses below it.
        File.WriteAllText(input, string.Join("\n", Enumerable.Range(0, 3000).Select(index =>
            "<p>A long source " + index.ToString(System.Globalization.CultureInfo.InvariantCulture) + " that exceeds the bounded PDF output.</p>")));
        var request = new OfficeBatchOptions {
            InputPaths = [input],
            InputDirectory = scope.Input,
            OutputDirectory = scope.Output,
            ResumeFrom = Path.Combine(scope.State, "plan.json"),
            MaximumOutputBytes = 16 * 1024,
            ConversionOptions = new() {
                Markdown = new() {
                    BaseDirectory = scope.Input,
                    ResourcePolicy = new PdfResourcePolicy { AllowLocalFileAccess = true }
                }
            }
        };
        Assert.Equal(1, (await OfficeBatchExecutor.RunAsync(new OfficeConversionExecutor(), request)).Failed);
        File.WriteAllText(input, "Corrected source");
        Assert.Equal(1, (await OfficeBatchExecutor.RunAsync(new OfficeConversionExecutor(), request)).Failed);
        var repaired = await OfficeBatchExecutor.RunAsync(new OfficeConversionExecutor(), request with { RetryFailed = true });
        Assert.Equal(0, repaired.Failed); Assert.Equal(1, repaired.Completed);
        Assert.Equal(1, (await OfficeBatchExecutor.RunAsync(new OfficeConversionExecutor(), request)).Reused);
    }

    private sealed class ChangeResource(string resource) : IOfficePublicationGuard {
        public Task<bool> CanPublishAsync(string path, bool directory, CancellationToken token) {
            if (!directory && Path.GetExtension(path) == ".pdf") File.WriteAllText(resource, "Changed at publication");
            return Task.FromResult(true);
        }
    }

    private sealed class Items : IProgress<OfficeBatchItemReport> {
        public ConcurrentBag<OfficeBatchItemReport> Values { get; } = [];
        public void Report(OfficeBatchItemReport item) => Values.Add(item);
    }
    private sealed class BatchDirectory : IDisposable {
        private readonly string _root = Path.Combine(Path.GetTempPath(), "officeimo-conversion-batch-" + Guid.NewGuid().ToString("N"));
        public string Input => Path.Combine(_root, "input");
        public string Output => Path.Combine(_root, "output");
        public string State => Path.Combine(_root, "state");
        public BatchDirectory() => Directory.CreateDirectory(Input);
        public void Dispose() => Directory.Delete(_root, recursive: true);
    }
}