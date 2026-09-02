using OfficeIMO.Provenance;
using OfficeIMO.Epub;
using OfficeIMO.Excel;
using OfficeIMO.OpenDocument;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.Visio;
using OfficeIMO.Word;
using System.Text;

namespace OfficeIMO.Workflows.Tests;

public sealed class OfficeProvenanceWorkflowTests {
    [Fact]
    public void CatalogRoutesFormatsToCanonicalOwners() {
        Assert.Equal("OfficeIMO.Word", OfficeProvenanceWorkflowCatalog.FindByPath("report.docx")?.OwnerPackage);
        Assert.Equal("OfficeIMO.Html", OfficeProvenanceWorkflowCatalog.FindByPath("page.HTML")?.OwnerPackage);
        Assert.Equal("OfficeIMO.Core", OfficeProvenanceWorkflowCatalog.FindByPath("image.png")?.OwnerPackage);
        Assert.Equal("OfficeIMO.Excel", OfficeProvenanceWorkflowCatalog.FindByPath("workbook.xlsb")?.OwnerPackage);
        Assert.Null(OfficeProvenanceWorkflowCatalog.FindByPath("archive.zip"));
        Assert.True(OfficeProvenanceWorkflowCatalog.All.Single(item => item.Id == "core-detected").CanRemove);
    }

    [Fact]
    public async Task InspectUsesHtmlOwnerAndReturnsStructuralEvidence() {
        using var scope = new TempScope();
        string input = scope.Write("page.html", HtmlWithExternalManifest("body"));

        OfficeProvenanceWorkflowResult result = await new OfficeWorkflowRunner().RunProvenanceAsync(
            new OfficeProvenanceWorkflowRequest {
                Operation = OfficeProvenanceWorkflowOperation.Inspect,
                InputPath = input
            });

        Assert.True(result.Succeeded, result.Summary);
        Assert.Equal("OfficeIMO.Html", result.OwnerPackage);
        Assert.Equal(OfficeProvenanceAssetFormat.Html, result.Inspection?.Format);
        Assert.Single(result.Inspection!.Evidence);
        Assert.Null(result.OutputPath);
    }

    [Fact]
    public async Task EmptyInputPathReturnsAStableValidationFailure() {
        OfficeProvenanceWorkflowResult result = await new OfficeWorkflowRunner().RunProvenanceAsync(
            new OfficeProvenanceWorkflowRequest {
                Operation = OfficeProvenanceWorkflowOperation.Inspect,
                InputPath = string.Empty
            });

        Assert.Equal(OfficeWorkflowStatus.Failed, result.Status);
        Assert.Equal(OfficeWorkflowFailureKind.ValidationFailed, result.FailureKind);
        Assert.Equal("OfficeIMO.Core", result.OwnerPackage);
    }

    [Fact]
    public async Task InspectAcceptsRepresentativeArtifactsFromEveryRegisteredOwner() {
        using var scope = new TempScope();
        string word = Path.Combine(scope.Path, "report.docx");
        using (WordDocument document = WordDocument.Create(word)) {
            document.AddParagraph("workflow provenance");
            document.Save();
        }
        string excel = Path.Combine(scope.Path, "workbook.xlsx");
        using (ExcelDocument document = ExcelDocument.Create(excel)) {
            document.AddWorksheet("Data").Cell(1, 1, "workflow provenance");
            document.Save();
        }
        string powerPoint = Path.Combine(scope.Path, "deck.pptx");
        using (PowerPointPresentation presentation = PowerPointPresentation.Create(powerPoint)) {
            presentation.AddSlide().AddTextBoxPoints("workflow provenance", 20, 20, 300, 60);
            presentation.Save();
        }
        string visio = Path.Combine(scope.Path, "drawing.vsdx");
        VisioDocument visioDocument = VisioDocument.Create(visio);
        visioDocument.AddPage("Page-1", 8.5, 11);
        visioDocument.Save();
        string odt = Path.Combine(scope.Path, "document.odt");
        OdtDocument openDocument = OdtDocument.Create();
        openDocument.AddParagraph("workflow provenance");
        openDocument.Save(odt);
        string epub = Path.Combine(scope.Path, "publication.epub");
        CreateEpub(epub);
        string pdf = Path.Combine(scope.Path, "document.pdf");
        PdfDocument.Create(document => document.Page(page => page.Content(content =>
            content.Item(item => item.Paragraph(paragraph => paragraph.Text("workflow provenance")))))).Save(pdf);
        string html = scope.Write("document.html", "<!doctype html><html><body>workflow provenance</body></html>");
        string markdown = scope.Write("document.md", "# Workflow provenance");
        string png = Path.Combine(scope.Path, "image.png");
        File.WriteAllBytes(png, Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="));

        var expected = new Dictionary<string, string>(StringComparer.Ordinal) {
            [word] = "OfficeIMO.Word",
            [excel] = "OfficeIMO.Excel",
            [powerPoint] = "OfficeIMO.PowerPoint",
            [visio] = "OfficeIMO.Visio",
            [odt] = "OfficeIMO.OpenDocument",
            [epub] = "OfficeIMO.Epub",
            [pdf] = "OfficeIMO.Pdf",
            [html] = "OfficeIMO.Html",
            [markdown] = "OfficeIMO.Markdown",
            [png] = "OfficeIMO.Core"
        };
        var runner = new OfficeWorkflowRunner();

        foreach ((string path, string owner) in expected) {
            OfficeProvenanceWorkflowResult result = await runner.RunProvenanceAsync(
                new OfficeProvenanceWorkflowRequest {
                    Operation = OfficeProvenanceWorkflowOperation.Inspect,
                    InputPath = path
                });
            Assert.True(result.Succeeded, path + ": " + result.Summary);
            Assert.Equal(owner, result.OwnerPackage);
        }
    }

    [Fact]
    public async Task AssessCombinesOwnerEvidenceTextIntegrityAndInjectedProviders() {
        using var scope = new TempScope();
        string input = scope.Write("page.html", HtmlWithExternalManifest("review \u202Ethis"));
        var runner = new OfficeWorkflowRunner(new TestVerifier(), [new TestDetector()]);

        OfficeProvenanceWorkflowResult result = await runner.RunProvenanceAsync(
            new OfficeProvenanceWorkflowRequest {
                Operation = OfficeProvenanceWorkflowOperation.Assess,
                InputPath = input
            });

        Assert.True(result.Succeeded, result.Summary);
        OfficeProvenanceAssessmentReport assessment = Assert.IsType<OfficeProvenanceAssessmentReport>(result.Assessment);
        Assert.Single(assessment.Structural.Evidence);
        Assert.Equal(OfficeProvenanceVerificationStatus.Valid, assessment.Verification?.Status);
        Assert.Contains(assessment.TextIntegrity!.Findings, finding => finding.Kind == OfficeTextIntegrityFindingKind.BidirectionalControl);
        Assert.Equal(OfficeProvenanceSignalStatus.Detected, Assert.Single(assessment.ProviderSignals).Status);
    }

    [Fact]
    public async Task RemoveStagesReopensAndPublishesThroughHtmlOwner() {
        using var scope = new TempScope();
        string input = scope.Write("page.html", HtmlWithExternalManifest("keep me"));
        string output = Path.Combine(scope.Path, "cleaned.html");

        OfficeProvenanceWorkflowResult result = await new OfficeWorkflowRunner().RunProvenanceAsync(
            new OfficeProvenanceWorkflowRequest {
                Operation = OfficeProvenanceWorkflowOperation.Remove,
                InputPath = input,
                OutputPath = output,
                ConflictPolicy = OfficeWorkflowConflictPolicy.Fail
            });

        Assert.True(result.Succeeded, result.Summary);
        Assert.Equal(Path.GetFullPath(output), result.OutputPath);
        Assert.True(result.WasChanged);
        Assert.Single(result.Before!.Evidence);
        Assert.Empty(result.After!.Evidence);
        Assert.Contains("keep me", File.ReadAllText(output), StringComparison.Ordinal);
        Assert.DoesNotContain("c2pa-manifest", File.ReadAllText(output), StringComparison.OrdinalIgnoreCase);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "ProvenanceOutputReopened");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "AtomicPublication");
        Assert.Empty(Directory.EnumerateFiles(scope.Path, ".cleaned.*.html"));
    }

    [Fact]
    public async Task FailedConflictPreservesExistingOutputAndCleansStaging() {
        using var scope = new TempScope();
        string input = scope.Write("page.html", HtmlWithExternalManifest("input"));
        string output = scope.Write("cleaned.html", "existing");

        OfficeProvenanceWorkflowResult result = await new OfficeWorkflowRunner().RunProvenanceAsync(
            new OfficeProvenanceWorkflowRequest {
                Operation = OfficeProvenanceWorkflowOperation.Remove,
                InputPath = input,
                OutputPath = output,
                ConflictPolicy = OfficeWorkflowConflictPolicy.Fail
            });

        Assert.Equal(OfficeWorkflowStatus.Failed, result.Status);
        Assert.Equal(OfficeWorkflowFailureKind.OutputFailed, result.FailureKind);
        Assert.Equal("existing", File.ReadAllText(output));
        Assert.Empty(Directory.EnumerateFiles(scope.Path, ".cleaned.*.html"));
    }

    [Fact]
    public async Task GenericZipCanBeInspectedButNotMutatedWithoutAFormatOwner() {
        using var scope = new TempScope();
        string input = Path.Combine(scope.Path, "archive.zip");
        using (System.IO.Compression.ZipArchive archive = System.IO.Compression.ZipFile.Open(input, System.IO.Compression.ZipArchiveMode.Create)) {
            using StreamWriter writer = new(archive.CreateEntry("file.txt").Open());
            writer.Write("content");
        }

        OfficeProvenanceWorkflowResult inspection = await new OfficeWorkflowRunner().RunProvenanceAsync(
            new OfficeProvenanceWorkflowRequest {
                Operation = OfficeProvenanceWorkflowOperation.Inspect,
                InputPath = input
            });
        OfficeProvenanceWorkflowResult removal = await new OfficeWorkflowRunner().RunProvenanceAsync(
            new OfficeProvenanceWorkflowRequest {
                Operation = OfficeProvenanceWorkflowOperation.Remove,
                InputPath = input
            });

        Assert.True(inspection.Succeeded, inspection.Summary);
        Assert.Equal(OfficeProvenanceAssetFormat.ZipPackage, inspection.Inspection?.Format);
        Assert.Equal(OfficeWorkflowFailureKind.UnsupportedInput, removal.FailureKind);
        Assert.False(File.Exists(Path.Combine(scope.Path, "archive.provenance-cleaned.zip")));
    }

    [Fact]
    public async Task BatchMaterializationIsBoundedBeforeExecution() {
        using var scope = new TempScope();
        string input = scope.Write("page.html", "<html><body>ok</body></html>");
        OfficeProvenanceWorkflowRequest[] requests = Enumerable.Range(0, 3)
            .Select(index => new OfficeProvenanceWorkflowRequest {
                Id = "request-" + index,
                Operation = OfficeProvenanceWorkflowOperation.Inspect,
                InputPath = input
            })
            .ToArray();

        ArgumentException exception = await Assert.ThrowsAsync<ArgumentException>(() =>
            new OfficeWorkflowRunner().RunProvenanceBatchAsync(
                requests,
                new OfficeProvenanceWorkflowBatchOptions { MaximumRequests = 2 }));

        Assert.Contains("limit of 2", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public async Task BatchRejectsDuplicateRemovalOutputsBeforeExecution() {
        using var scope = new TempScope();
        string first = scope.Write("first.html", HtmlWithExternalManifest("first"));
        string second = scope.Write("second.html", HtmlWithExternalManifest("second"));
        string output = Path.Combine(scope.Path, "cleaned.html");
        OfficeProvenanceWorkflowRequest[] requests = new[] { first, second }
            .Select((input, index) => new OfficeProvenanceWorkflowRequest {
                Id = "request-" + index,
                Operation = OfficeProvenanceWorkflowOperation.Remove,
                InputPath = input,
                OutputPath = output,
                ConflictPolicy = OfficeWorkflowConflictPolicy.Replace
            })
            .ToArray();

        ArgumentException exception = await Assert.ThrowsAsync<ArgumentException>(() =>
            new OfficeWorkflowRunner().RunProvenanceBatchAsync(requests));

        Assert.Contains("same output path", exception.Message, StringComparison.Ordinal);
        Assert.False(File.Exists(output));
    }

    [Fact]
    public async Task BatchRejectsRemovalOutputThatOverlapsAnotherInputBeforeExecution() {
        using var scope = new TempScope();
        string first = scope.Write("first.html", HtmlWithExternalManifest("first"));
        string second = scope.Write("second.html", HtmlWithExternalManifest("second"));
        string originalSecond = File.ReadAllText(second);

        ArgumentException exception = await Assert.ThrowsAsync<ArgumentException>(() =>
            new OfficeWorkflowRunner().RunProvenanceBatchAsync([
                new OfficeProvenanceWorkflowRequest {
                    Id = "first",
                    Operation = OfficeProvenanceWorkflowOperation.Remove,
                    InputPath = first,
                    OutputPath = second,
                    ConflictPolicy = OfficeWorkflowConflictPolicy.Replace
                },
                new OfficeProvenanceWorkflowRequest {
                    Id = "second",
                    Operation = OfficeProvenanceWorkflowOperation.Inspect,
                    InputPath = second
                }
            ]));

        Assert.Contains("overlaps another batch request's input", exception.Message, StringComparison.Ordinal);
        Assert.Equal(originalSecond, File.ReadAllText(second));
    }

    [Fact]
    public async Task RemovalPreflightUsesRemovalLimitsInsteadOfInspectionLimits() {
        using var scope = new TempScope();
        string input = scope.Write("page.html", HtmlWithExternalManifest("body"));
        string output = Path.Combine(scope.Path, "cleaned.html");
        var request = new OfficeProvenanceWorkflowRequest {
            Operation = OfficeProvenanceWorkflowOperation.Remove,
            InputPath = input,
            OutputPath = output
        };
        request.Inspection.MaxAssetBytes = 1;

        OfficeProvenanceWorkflowResult result = await new OfficeWorkflowRunner().RunProvenanceAsync(request);

        Assert.True(result.Succeeded, result.Summary);
        Assert.True(File.Exists(output));
    }

    [Fact]
    public async Task PreCancelledBatchReturnsOneExplicitCancelledResult() {
        using var scope = new TempScope();
        string first = scope.Write("first.html", "<html><body>first</body></html>");
        string second = scope.Write("second.html", "<html><body>second</body></html>");
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        IReadOnlyList<OfficeProvenanceWorkflowResult> results = await new OfficeWorkflowRunner().RunProvenanceBatchAsync(
            [
                new OfficeProvenanceWorkflowRequest { Id = "first", Operation = OfficeProvenanceWorkflowOperation.Inspect, InputPath = first },
                new OfficeProvenanceWorkflowRequest { Id = "second", Operation = OfficeProvenanceWorkflowOperation.Inspect, InputPath = second }
            ],
            cancellationToken: cancellation.Token);

        OfficeProvenanceWorkflowResult result = Assert.Single(results);
        Assert.Equal("first", result.RequestId);
        Assert.Equal(OfficeWorkflowStatus.Cancelled, result.Status);
    }

    [Fact]
    public async Task CancellationBetweenSuccessfulItemsAddsAnExplicitCancelledResult() {
        using var scope = new TempScope();
        string first = scope.Write("first.html", "<html><body>first</body></html>");
        string second = scope.Write("second.html", "<html><body>second</body></html>");
        using var cancellation = new CancellationTokenSource();
        var progress = new CancellingProgress("first", cancellation);

        IReadOnlyList<OfficeProvenanceWorkflowResult> results = await new OfficeWorkflowRunner().RunProvenanceBatchAsync(
            [
                new OfficeProvenanceWorkflowRequest { Id = "first", Operation = OfficeProvenanceWorkflowOperation.Inspect, InputPath = first },
                new OfficeProvenanceWorkflowRequest { Id = "second", Operation = OfficeProvenanceWorkflowOperation.Inspect, InputPath = second }
            ],
            progress: progress,
            cancellationToken: cancellation.Token);

        Assert.Collection(
            results,
            result => Assert.Equal(OfficeWorkflowStatus.Completed, result.Status),
            result => Assert.Equal(OfficeWorkflowStatus.Cancelled, result.Status));
    }

    [Fact]
    public async Task BatchKeepsPerRequestFailuresWhenContinueOnFailureIsEnabled() {
        using var scope = new TempScope();
        string input = scope.Write("first.html", HtmlWithExternalManifest("first"));
        string missing = Path.Combine(scope.Path, "missing.html");
        string firstOutput = Path.Combine(scope.Path, "first-cleaned.html");
        string missingOutput = Path.Combine(scope.Path, "missing-cleaned.html");

        IReadOnlyList<OfficeProvenanceWorkflowResult> results = await new OfficeWorkflowRunner().RunProvenanceBatchAsync([
            new OfficeProvenanceWorkflowRequest {
                Id = "valid",
                Operation = OfficeProvenanceWorkflowOperation.Remove,
                InputPath = input,
                OutputPath = firstOutput
            },
            new OfficeProvenanceWorkflowRequest {
                Id = "missing",
                Operation = OfficeProvenanceWorkflowOperation.Remove,
                InputPath = missing,
                OutputPath = missingOutput
            }
        ]);

        Assert.Collection(
            results,
            result => Assert.Equal(OfficeWorkflowStatus.Completed, result.Status),
            result => Assert.Equal(OfficeWorkflowFailureKind.InputNotFound, result.FailureKind));
        Assert.True(File.Exists(firstOutput));
        Assert.False(File.Exists(missingOutput));
    }

    [Fact]
    public async Task AssessmentCancellationRaisedBySynchronousDetectorIsNotReportedAsCompleted() {
        using var scope = new TempScope();
        string input = scope.Write("page.html", "<html><body>body</body></html>");
        using var cancellation = new CancellationTokenSource();
        var runner = new OfficeWorkflowRunner(null, [new CancellingDetector(cancellation)]);

        OfficeProvenanceWorkflowResult result = await runner.RunProvenanceAsync(
            new OfficeProvenanceWorkflowRequest {
                Operation = OfficeProvenanceWorkflowOperation.Assess,
                InputPath = input
            },
            cancellationToken: cancellation.Token);

        Assert.Equal(OfficeWorkflowStatus.Cancelled, result.Status);
    }

    [Fact]
    public async Task SmallWorkflowLimitsRemainValidForTinyInputsAndOutputs() {
        using var scope = new TempScope();
        string input = scope.Write("page.html", HtmlWithExternalManifest("tiny"));
        string output = Path.Combine(scope.Path, "cleaned.html");
        var limits = new OfficeWorkflowLimits {
            MaximumInputBytes = 1024,
            MaximumOutputBytes = 1024
        };

        OfficeProvenanceWorkflowResult inspection = await new OfficeWorkflowRunner().RunProvenanceAsync(
            new OfficeProvenanceWorkflowRequest {
                Operation = OfficeProvenanceWorkflowOperation.Inspect,
                InputPath = input,
                Limits = limits
            });
        OfficeProvenanceWorkflowResult removal = await new OfficeWorkflowRunner().RunProvenanceAsync(
            new OfficeProvenanceWorkflowRequest {
                Operation = OfficeProvenanceWorkflowOperation.Remove,
                InputPath = input,
                OutputPath = output,
                Limits = limits
            });

        Assert.True(inspection.Succeeded, inspection.Summary);
        Assert.True(removal.Succeeded, removal.Summary);
        Assert.True(File.Exists(output));
    }

    [Fact]
    public async Task StagedOutputLimitFailureIsClassifiedAsOutputFailure() {
        using var scope = new TempScope();
        string input = scope.Write("page.html", HtmlWithExternalManifest("body"));
        string output = Path.Combine(scope.Path, "cleaned.html");

        OfficeProvenanceWorkflowResult result = await new OfficeWorkflowRunner().RunProvenanceAsync(
            new OfficeProvenanceWorkflowRequest {
                Operation = OfficeProvenanceWorkflowOperation.Remove,
                InputPath = input,
                OutputPath = output,
                Limits = new OfficeWorkflowLimits { MaximumOutputBytes = 1 }
            });

        Assert.Equal(OfficeWorkflowStatus.Failed, result.Status);
        Assert.Equal(OfficeWorkflowFailureKind.OutputFailed, result.FailureKind);
        Assert.False(File.Exists(output));
        Assert.Empty(Directory.EnumerateFiles(scope.Path, ".cleaned.*.html"));
    }

    [Fact]
    public async Task SignatureDetectedImageWithUnknownExtensionCanBeRemoved() {
        using var scope = new TempScope();
        string input = Path.Combine(scope.Path, "image.bin");
        string output = Path.Combine(scope.Path, "cleaned.bin");
        File.WriteAllBytes(input, Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="));

        OfficeProvenanceWorkflowResult result = await new OfficeWorkflowRunner().RunProvenanceAsync(
            new OfficeProvenanceWorkflowRequest {
                Operation = OfficeProvenanceWorkflowOperation.Remove,
                InputPath = input,
                OutputPath = output
            });

        Assert.True(result.Succeeded, result.Summary);
        Assert.Equal(OfficeProvenanceAssetFormat.Png, result.Before?.Format);
        Assert.Equal(OfficeProvenanceAssetFormat.Png, result.After?.Format);
        Assert.True(File.Exists(output));
    }

    [Fact]
    public async Task PreCancelledRemovalPublishesNothing() {
        using var scope = new TempScope();
        string input = scope.Write("page.html", HtmlWithExternalManifest("body"));
        string output = Path.Combine(scope.Path, "cleaned.html");
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        OfficeProvenanceWorkflowResult result = await new OfficeWorkflowRunner().RunProvenanceAsync(
            new OfficeProvenanceWorkflowRequest {
                Operation = OfficeProvenanceWorkflowOperation.Remove,
                InputPath = input,
                OutputPath = output
            }, cancellationToken: cancellation.Token);

        Assert.Equal(OfficeWorkflowStatus.Cancelled, result.Status);
        Assert.False(File.Exists(output));
    }

    private static string HtmlWithExternalManifest(string body) =>
        "<!doctype html><html><head><link rel=\"c2pa-manifest\" href=\"claim.c2pa\"></head><body>" + body + "</body></html>";

    private static void CreateEpub(string path) {
        using var stream = File.Create(path);
        using var archive = new System.IO.Compression.ZipArchive(
            stream,
            System.IO.Compression.ZipArchiveMode.Create,
            leaveOpen: false);
        WriteEntry(archive, "mimetype", "application/epub+zip", System.IO.Compression.CompressionLevel.NoCompression);
        WriteEntry(
            archive,
            "META-INF/container.xml",
            "<?xml version=\"1.0\"?><container xmlns=\"urn:oasis:names:tc:opendocument:xmlns:container\"><rootfiles><rootfile full-path=\"content.opf\" media-type=\"application/oebps-package+xml\"/></rootfiles></container>");
        WriteEntry(
            archive,
            "content.opf",
            "<?xml version=\"1.0\"?><package xmlns=\"http://www.idpf.org/2007/opf\" version=\"3.0\" unique-identifier=\"id\"><metadata xmlns:dc=\"http://purl.org/dc/elements/1.1/\"><dc:identifier id=\"id\">urn:uuid:" + Guid.NewGuid() + "</dc:identifier></metadata><manifest/><spine/></package>");
    }

    private static void WriteEntry(
        System.IO.Compression.ZipArchive archive,
        string name,
        string value,
        System.IO.Compression.CompressionLevel compression = System.IO.Compression.CompressionLevel.Optimal) {
        System.IO.Compression.ZipArchiveEntry entry = archive.CreateEntry(name, compression);
        using var writer = new StreamWriter(
            entry.Open(),
            new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        writer.Write(value);
    }

    private sealed class TestVerifier : IOfficeProvenanceVerifier {
        public string Name => "test-verifier";

        public OfficeProvenanceVerificationResult Verify(
            string filePath,
            OfficeProvenanceVerificationOptions? options = null) =>
            new(OfficeProvenanceVerificationStatus.Valid, Name, ["content binding verified"]);
    }

    private sealed class TestDetector : IOfficeProvenanceSignalDetector {
        public string Name => "test-detector";
        public OfficeProvenanceSignalKind SignalKind => OfficeProvenanceSignalKind.DeterministicArtifact;

        public OfficeProvenanceSignalResult Detect(string filePath) =>
            new(Name, SignalKind, OfficeProvenanceSignalStatus.Detected, ["test signal"]);
    }

    private sealed class CancellingDetector : IOfficeProvenanceSignalDetector {
        private readonly CancellationTokenSource _cancellation;

        internal CancellingDetector(CancellationTokenSource cancellation) => _cancellation = cancellation;

        public string Name => "cancelling-detector";
        public OfficeProvenanceSignalKind SignalKind => OfficeProvenanceSignalKind.DeterministicArtifact;

        public OfficeProvenanceSignalResult Detect(string filePath) {
            _cancellation.Cancel();
            return new OfficeProvenanceSignalResult(Name, SignalKind, OfficeProvenanceSignalStatus.NotDetected);
        }
    }

    private sealed class CancellingProgress : IProgress<OfficeWorkflowProgress> {
        private readonly string _requestId;
        private readonly CancellationTokenSource _cancellation;

        internal CancellingProgress(string requestId, CancellationTokenSource cancellation) {
            _requestId = requestId;
            _cancellation = cancellation;
        }

        public void Report(OfficeWorkflowProgress value) {
            if (value.RequestId == _requestId && value.Stage == "complete") _cancellation.Cancel();
        }
    }

    private sealed class TempScope : IDisposable {
        internal TempScope() {
            Path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "OfficeIMO-ProvenanceWorkflow-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(Path);
        }

        internal string Path { get; }

        internal string Write(string fileName, string contents) {
            string path = System.IO.Path.Combine(Path, fileName);
            File.WriteAllText(path, contents, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
            return path;
        }

        public void Dispose() {
            try {
                if (Directory.Exists(Path)) Directory.Delete(Path, recursive: true);
            } catch (IOException) {
                // Best-effort test cleanup.
            } catch (UnauthorizedAccessException) {
                // Best-effort test cleanup.
            }
        }
    }
}
