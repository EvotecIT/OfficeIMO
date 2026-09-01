using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Excel;
using OfficeIMO.Html;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;
using OfficeIMO.Workflows;

namespace OfficeIMO.Workflows.Tests;

public sealed class OfficeWorkflowRunnerTests {
    [Fact]
    public void PdfLoadOptionsUseTheValidatedWorkflowByteBudget() {
        const long budget = 768L * 1024L * 1024L;

        PdfLoadOptions options = OfficeWorkflowRunner.CreatePdfLoadOptions("open", budget);

        Assert.Equal("open", options.Password);
        Assert.Equal(budget, options.Limits.MaxInputBytes);
    }

    [Fact]
    public void CatalogProjectsExactlyTheEightOfficePdfRoutes() {
        Assert.Equal(8, OfficeWorkflowCatalog.Routes.Count);
        Assert.Equal(
            ["docx-pdf", "html-pdf", "pdf-docx", "pdf-html", "pdf-pptx", "pdf-xlsx", "pptx-pdf", "xlsx-pdf"],
            OfficeWorkflowCatalog.Routes.Select(route => route.Id).OrderBy(id => id, StringComparer.Ordinal));
        Assert.All(OfficeWorkflowCatalog.Routes, route => Assert.StartsWith("OfficeIMO.", route.Engine, StringComparison.Ordinal));
    }

    [Theory]
    [InlineData("docx-pdf")]
    [InlineData("xlsx-pdf")]
    [InlineData("pptx-pdf")]
    [InlineData("html-pdf")]
    [InlineData("pdf-docx")]
    [InlineData("pdf-xlsx")]
    [InlineData("pdf-pptx")]
    [InlineData("pdf-html")]
    public async Task EveryCatalogRoutePublishesAnArtifactThatWasReopened(string routeId) {
        using var scope = new TestDirectory();
        OfficeWorkflowRoute route = Assert.Single(OfficeWorkflowCatalog.Routes, item => item.Id == routeId);
        string input = CreateInput(scope.Path, routeId);
        string output = System.IO.Path.Combine(scope.Path, "result" + NormalizeExtension(route.TargetExtension));
        var runner = new OfficeWorkflowRunner();

        OfficeWorkflowResult result = await runner.RunAsync(new OfficeWorkflowRequest {
            Operation = OfficeWorkflowOperation.Convert,
            InputPath = input,
            OutputPath = output,
            ConversionRouteId = routeId,
            ConflictPolicy = OfficeWorkflowConflictPolicy.Fail
        });

        Assert.True(result.Succeeded, result.Summary);
        Assert.Equal(OfficeWorkflowFailureKind.None, result.FailureKind);
        Assert.Equal(output, result.OutputPath);
        Assert.True(File.Exists(output));
        Assert.True(result.OutputBytes > 0);
        OfficeWorkflowDiagnostic reopened = Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "OutputReopened");
        Assert.Equal(result.OutputBytes.ToString(System.Globalization.CultureInfo.InvariantCulture), reopened.Details["stagedBytes"]);
        Assert.Equal(NormalizeExtension(route.TargetExtension), reopened.Details["format"]);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "AtomicPublication");
    }

    [Fact]
    public async Task ConversionStopsWhileSerializingAtTheConfiguredOutputLimit() {
        using var scope = new TestDirectory();
        string input = CreateInput(scope.Path, "docx-pdf");
        string output = System.IO.Path.Combine(scope.Path, "bounded.pdf");

        OfficeWorkflowResult result = await new OfficeWorkflowRunner().RunAsync(new OfficeWorkflowRequest {
            Operation = OfficeWorkflowOperation.Convert,
            ConversionRouteId = "docx-pdf",
            InputPath = input,
            OutputPath = output,
            ConflictPolicy = OfficeWorkflowConflictPolicy.Fail,
            Limits = new OfficeWorkflowLimits {
                MaximumInputBytes = 16L * 1024L * 1024L,
                MaximumOutputBytes = 128L
            }
        });

        Assert.Equal(OfficeWorkflowStatus.Failed, result.Status);
        Assert.Equal(OfficeWorkflowFailureKind.OperationFailed, result.FailureKind);
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Message.Contains("while it was being serialized", StringComparison.Ordinal));
        Assert.False(File.Exists(output));
    }

    [Fact]
    public async Task HtmlToPdfRejectsOutputProfilesItCannotHonor() {
        using var scope = new TestDirectory();
        string input = CreateInput(scope.Path, "html-pdf");

        OfficeWorkflowResult result = await new OfficeWorkflowRunner().RunAsync(new OfficeWorkflowRequest {
            Operation = OfficeWorkflowOperation.Convert,
            ConversionRouteId = "html-pdf",
            InputPath = input,
            OutputPath = System.IO.Path.Combine(scope.Path, "lightweight.pdf"),
            OutputProfile = OfficeWorkflowOutputProfile.Lightweight
        });

        Assert.Equal(OfficeWorkflowStatus.Failed, result.Status);
        Assert.Equal(OfficeWorkflowFailureKind.ValidationFailed, result.FailureKind);
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Message.Contains("supports only the Faithful output profile", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData("pdf-docx")]
    [InlineData("pdf-xlsx")]
    [InlineData("pdf-pptx")]
    [InlineData("pdf-html")]
    public async Task PdfImportRoutesRejectOutputProfilesTheyCannotHonor(string routeId) {
        using var scope = new TestDirectory();
        OfficeWorkflowRoute route = Assert.Single(OfficeWorkflowCatalog.Routes, item => item.Id == routeId);
        string input = CreateInput(scope.Path, routeId);
        string output = System.IO.Path.Combine(scope.Path, "unsupported" + NormalizeExtension(route.TargetExtension));

        OfficeWorkflowResult result = await new OfficeWorkflowRunner().RunAsync(new OfficeWorkflowRequest {
            Operation = OfficeWorkflowOperation.Convert,
            ConversionRouteId = routeId,
            InputPath = input,
            OutputPath = output,
            OutputProfile = OfficeWorkflowOutputProfile.Lightweight
        });

        Assert.Equal(OfficeWorkflowStatus.Failed, result.Status);
        Assert.Contains("supports only the Faithful output profile", result.Summary, StringComparison.Ordinal);
        Assert.False(File.Exists(output));
    }

    [Fact]
    public async Task OptimizeRejectsTextOnlyProfileInsteadOfSubstitutingLosslessWebProfile() {
        using var scope = new TestDirectory();
        string input = CreatePdf(scope.Path, "source.pdf");
        string output = Path.Combine(scope.Path, "text-only.pdf");

        OfficeWorkflowResult result = await new OfficeWorkflowRunner().RunAsync(new OfficeWorkflowRequest {
            Operation = OfficeWorkflowOperation.Optimize,
            InputPath = input,
            OutputPath = output,
            OutputProfile = OfficeWorkflowOutputProfile.TextOnly
        });

        Assert.Equal(OfficeWorkflowStatus.Failed, result.Status);
        Assert.Contains("does not support the TextOnly output profile", result.Summary, StringComparison.Ordinal);
        Assert.False(File.Exists(output));
    }

    [Fact]
    public void InputReaderEnforcesTheLimitAgainstBytesReadFromTheOpenedHandle() {
        using var source = new UnderreportedLengthStream(new byte[17], reportedLength: 16);

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            OfficeWorkflowInputReader.ReadAllBytes(source, "growing.pdf", 16, CancellationToken.None));

        Assert.Contains("above the configured 16-byte limit", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlInputPreservesTheSourceFileAsItsBaseUri() {
        using var scope = new TestDirectory();
        string input = Path.Combine(scope.Path, "source.html");
        byte[] html = Encoding.UTF8.GetBytes("<!doctype html><html><body><img src='assets/chart.png'></body></html>");

        HtmlConversionDocument document = OfficeWorkflowRunner.ParseHtmlInput(html, input);

        Assert.Equal(new Uri(Path.GetFullPath(input)), document.BaseUri);
    }

    [Fact]
    public async Task CancellationDuringActiveHtmlConversionStopsBeforePublication() {
        using var scope = new TestDirectory();
        string input = Path.Combine(scope.Path, "source.html");
        string output = Path.Combine(scope.Path, "cancelled.pdf");
        var html = new StringBuilder("<!doctype html><html><body>");
        for (int index = 0; index < 40_000; index++) {
            html.Append("<p>Cancellation checkpoint ").Append(index).Append(" with enough text to exercise layout.</p>");
        }
        html.Append("</body></html>");
        await File.WriteAllTextAsync(input, html.ToString());
        using var cancellation = new CancellationTokenSource();
        var executionReported = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var progress = new InlineProgress<OfficeWorkflowProgress>(update => {
            if (update.Stage == "execute") executionReported.TrySetResult();
        });

        Task<OfficeWorkflowResult> run = new OfficeWorkflowRunner().RunAsync(new OfficeWorkflowRequest {
            Operation = OfficeWorkflowOperation.Convert,
            InputPath = input,
            OutputPath = output,
            ConversionRouteId = "html-pdf"
        }, progress, cancellation.Token);
        await executionReported.Task.WaitAsync(TimeSpan.FromSeconds(5));
        await Task.Delay(25);
        Assert.False(run.IsCompleted);
        cancellation.Cancel();

        OfficeWorkflowResult result = await run;

        Assert.True(result.Status == OfficeWorkflowStatus.Cancelled, result.Summary);
        Assert.False(File.Exists(output));
        Assert.Empty(Directory.GetFiles(scope.Path, ".*.tmp"));
    }

    [Fact]
    public async Task BatchProgressUsesStableIndicesAndReturnsBeforeTheFinalCallback() {
        using var scope = new TestDirectory();
        string first = CreatePdf(scope.Path, "first.pdf");
        string second = CreatePdf(scope.Path, "second.pdf");
        var updates = new List<OfficeWorkflowProgress>();

        IReadOnlyList<OfficeWorkflowResult> results = await new OfficeWorkflowRunner().RunBatchAsync([
            new OfficeWorkflowRequest { Id = "first", Operation = OfficeWorkflowOperation.Inspect, InputPath = first },
            new OfficeWorkflowRequest { Id = "second", Operation = OfficeWorkflowOperation.Inspect, InputPath = second }
        ], new InlineProgress<OfficeWorkflowProgress>(updates.Add));

        Assert.Equal(2, results.Count);
        Assert.NotEmpty(updates);
        Assert.All(updates.Where(update => update.RequestId == "first"), update => Assert.StartsWith("1 of 2", update.Message, StringComparison.Ordinal));
        Assert.All(updates.Where(update => update.RequestId == "second"), update => Assert.StartsWith("2 of 2", update.Message, StringComparison.Ordinal));
        Assert.Equal(1D, updates[^1].OverallFraction);
        Assert.True(updates.Select(update => update.OverallFraction).SequenceEqual(updates.Select(update => update.OverallFraction).OrderBy(value => value)));
    }

    [Fact]
    public async Task PreCancelledRequestPublishesNothingAndLeavesNoStagingFile() {
        using var scope = new TestDirectory();
        string input = CreatePdf(scope.Path, "source.pdf");
        string output = System.IO.Path.Combine(scope.Path, "cancelled.pdf");
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        OfficeWorkflowResult result = await new OfficeWorkflowRunner().RunAsync(new OfficeWorkflowRequest {
            Operation = OfficeWorkflowOperation.Optimize,
            InputPath = input,
            OutputPath = output
        }, cancellationToken: cancellation.Token);

        Assert.Equal(OfficeWorkflowStatus.Cancelled, result.Status);
        Assert.Equal(OfficeWorkflowFailureKind.None, result.FailureKind);
        Assert.False(File.Exists(output));
        Assert.Empty(Directory.GetFiles(scope.Path, ".*.tmp"));
    }

    [Fact]
    public async Task CollisionPoliciesFailRenameAndReplaceWithoutPartialArtifacts() {
        using var scope = new TestDirectory();
        string input = CreatePdf(scope.Path, "source.pdf");
        string output = System.IO.Path.Combine(scope.Path, "optimized.pdf");
        await File.WriteAllTextAsync(output, "existing");
        var runner = new OfficeWorkflowRunner();

        OfficeWorkflowResult failed = await runner.RunAsync(Optimize(input, output, OfficeWorkflowConflictPolicy.Fail));
        Assert.Equal(OfficeWorkflowStatus.Failed, failed.Status);
        Assert.Equal(OfficeWorkflowFailureKind.OutputFailed, failed.FailureKind);
        Assert.Equal("existing", await File.ReadAllTextAsync(output));

        OfficeWorkflowResult renamed = await runner.RunAsync(Optimize(input, output, OfficeWorkflowConflictPolicy.Rename));
        Assert.True(renamed.Succeeded, renamed.Summary);
        Assert.Equal(System.IO.Path.Combine(scope.Path, "optimized (1).pdf"), renamed.OutputPath);
        Assert.Equal("existing", await File.ReadAllTextAsync(output));

        OfficeWorkflowResult replaced = await runner.RunAsync(Optimize(input, output, OfficeWorkflowConflictPolicy.Replace));
        Assert.True(replaced.Succeeded, replaced.Summary);
        Assert.NotEqual("existing", await File.ReadAllTextAsync(output));
        Assert.Empty(Directory.GetFiles(scope.Path, ".*.tmp"));
    }

    [Fact]
    public async Task GeneralWorkflowResultClassifiesMissingInput() {
        using var scope = new TestDirectory();
        string missing = Path.Combine(scope.Path, "missing.pdf");
        OfficeWorkflowResult missingResult = await new OfficeWorkflowRunner().RunAsync(new OfficeWorkflowRequest {
            Operation = OfficeWorkflowOperation.Inspect,
            InputPath = missing
        });

        Assert.Equal(OfficeWorkflowFailureKind.InputNotFound, missingResult.FailureKind);
    }

    [Fact]
    public async Task RenamePolicySkipsARequestedPathOccupiedByADirectory() {
        using var scope = new TestDirectory();
        string input = CreatePdf(scope.Path, "source.pdf");
        string output = System.IO.Path.Combine(scope.Path, "optimized.pdf");
        Directory.CreateDirectory(output);

        OfficeWorkflowResult result = await new OfficeWorkflowRunner().RunAsync(
            Optimize(input, output, OfficeWorkflowConflictPolicy.Rename));

        Assert.True(result.Succeeded, result.Summary);
        Assert.Equal(System.IO.Path.Combine(scope.Path, "optimized (1).pdf"), result.OutputPath);
        Assert.True(Directory.Exists(output));
        Assert.True(File.Exists(result.OutputPath));
    }

    [Fact]
    public async Task MalformedPdfInspectionReturnsExplicitReadAndRepairEvidence() {
        using var scope = new TestDirectory();
        string validPath = CreatePdf(scope.Path, "valid.pdf");
        string malformedPath = System.IO.Path.Combine(scope.Path, "malformed.pdf");
        string raw = Encoding.Latin1.GetString(await File.ReadAllBytesAsync(validPath));
        string malformed = Regex.Replace(raw, "startxref\\r?\\n\\d+", "startxref\n0", RegexOptions.CultureInvariant);
        await File.WriteAllBytesAsync(malformedPath, Encoding.Latin1.GetBytes(malformed));

        OfficeWorkflowResult result = await new OfficeWorkflowRunner().RunAsync(new OfficeWorkflowRequest {
            Operation = OfficeWorkflowOperation.Inspect,
            InputPath = malformedPath
        });

        Assert.True(result.Succeeded, result.Summary);
        PdfHealthSnapshot before = Assert.IsType<PdfHealthSnapshot>(result.HealthReport?.Before);
        Assert.True(before.CanRead);
        Assert.True(before.RepairCount > 0 || before.Diagnostics.Count > 0);
    }

    [Fact]
    public async Task MalformedRepairPlanUsesCanonicalPlannerAndMatchesExecution() {
        using var scope = new TestDirectory();
        string validPath = CreatePdf(scope.Path, "valid.pdf");
        string malformedPath = Path.Combine(scope.Path, "malformed.pdf");
        string raw = Encoding.Latin1.GetString(await File.ReadAllBytesAsync(validPath));
        string malformed = Regex.Replace(raw, "startxref\\r?\\n\\d+", "startxref\n0", RegexOptions.CultureInvariant);
        await File.WriteAllBytesAsync(malformedPath, Encoding.Latin1.GetBytes(malformed));
        string repairedPath = Path.Combine(scope.Path, "repaired.pdf");
        var runner = new OfficeWorkflowRunner();

        OfficeWorkflowResult plan = await runner.RunAsync(new OfficeWorkflowRequest {
            Operation = OfficeWorkflowOperation.RepairPlan,
            InputPath = malformedPath
        });
        OfficeWorkflowResult repair = await runner.RunAsync(new OfficeWorkflowRequest {
            Operation = OfficeWorkflowOperation.Repair,
            InputPath = malformedPath,
            OutputPath = repairedPath,
            ConflictPolicy = OfficeWorkflowConflictPolicy.Fail
        });

        Assert.True(plan.Succeeded, plan.Summary);
        Assert.True(plan.HealthReport!.Verified, plan.Summary);
        Assert.Equal("True", plan.HealthReport.Metrics["canCreateRepairArtifact"]);
        Assert.Equal("FullRewrite", plan.HealthReport.Metrics["canonicalMutationMode"]);
        Assert.True(repair.Succeeded, repair.Summary);
        Assert.True(repair.HealthReport!.Verified);
        Assert.True(File.Exists(repairedPath));
    }

    [Fact]
    public async Task EncryptedPdfInspectionIsReportableWithoutLeakingPassword() {
        using var scope = new TestDirectory();
        string path = System.IO.Path.Combine(scope.Path, "encrypted.pdf");
        byte[] encrypted = PdfDocument.Create(
            compose => compose.Page(page => page.Content(content => content.Item(item => item.Paragraph(paragraph => paragraph.Text("Protected local content"))))),
            new PdfOptions().SetEncryption("open", "owner"))
            .ToBytes();
        await File.WriteAllBytesAsync(path, encrypted);
        var runner = new OfficeWorkflowRunner();

        OfficeWorkflowResult blocked = await runner.RunAsync(new OfficeWorkflowRequest {
            Operation = OfficeWorkflowOperation.Inspect,
            InputPath = path
        });
        OfficeWorkflowResult opened = await runner.RunAsync(new OfficeWorkflowRequest {
            Operation = OfficeWorkflowOperation.Inspect,
            InputPath = path,
            PdfPassword = "open"
        });

        Assert.True(blocked.Succeeded);
        Assert.True(blocked.HealthReport!.Before.HasEncryption);
        Assert.False(blocked.HealthReport.Before.CanRead);
        Assert.True(opened.Succeeded);
        Assert.True(opened.HealthReport!.Before.CanRead);
        Assert.DoesNotContain(opened.Diagnostics, diagnostic => diagnostic.Message.Contains("open", StringComparison.Ordinal));
    }

    [Fact]
    public async Task TaggedPdfInspectionPreservesTaggedStateAsTypedEvidence() {
        using var scope = new TestDirectory();
        string path = System.IO.Path.Combine(scope.Path, "tagged.pdf");
        byte[] tagged = PdfDocument.Create(
            compose => compose.Page(page => page.Content(content => content.Item(item => item.Paragraph(paragraph => paragraph.Text("Tagged document"))))),
            new PdfOptions().ConfigurePdfAGroundwork(PdfComplianceProfile.PdfA3A, "en-US"))
            .ToBytes();
        await File.WriteAllBytesAsync(path, tagged);

        OfficeWorkflowResult result = await new OfficeWorkflowRunner().RunAsync(new OfficeWorkflowRequest {
            Operation = OfficeWorkflowOperation.Inspect,
            InputPath = path
        });

        Assert.True(result.Succeeded);
        Assert.True(result.HealthReport!.Before.HasTaggedContent);
    }

    [Fact]
    public async Task SignedRewriteFailureLeavesExistingDestinationUntouched() {
        using var scope = new TestDirectory();
        string path = System.IO.Path.Combine(scope.Path, "signed.pdf");
        byte[] unsigned = PdfDocument.Create(compose => compose.Page(page => page.Content(content => content.Item(item => item.Paragraph(paragraph => paragraph.Text("Signed source")))))).ToBytes();
        PdfExternalSignaturePreparation preparation = PdfDocument.Load(unsigned).Security.PrepareExternalSignature(
            new PdfExternalSignatureOptions { FieldName = "Approval", ReservedSignatureContentsBytes = 512 });
        byte[] signed = preparation.Complete([0x30, 0x01, 0x00]).ToBytes();
        await File.WriteAllBytesAsync(path, signed);
        string output = System.IO.Path.Combine(scope.Path, "must-survive.pdf");
        byte[] sentinel = Encoding.UTF8.GetBytes("existing destination");
        await File.WriteAllBytesAsync(output, sentinel);

        OfficeWorkflowResult result = await new OfficeWorkflowRunner().RunAsync(
            Optimize(path, output, OfficeWorkflowConflictPolicy.Replace));

        Assert.Equal(OfficeWorkflowStatus.Failed, result.Status);
        Assert.Equal(sentinel, await File.ReadAllBytesAsync(output));
        Assert.Empty(Directory.GetFiles(scope.Path, ".*.tmp"));
    }

    [Fact]
    public async Task CompareAndSanitizeReturnExplicitBeforeAfterReports() {
        using var scope = new TestDirectory();
        string left = CreatePdf(scope.Path, "left.pdf", "Same text");
        string right = CreatePdf(scope.Path, "right.pdf", "Same text");
        string gallery = System.IO.Path.Combine(scope.Path, "comparison.html");
        var runner = new OfficeWorkflowRunner();

        OfficeWorkflowResult comparison = await runner.RunAsync(new OfficeWorkflowRequest {
            Operation = OfficeWorkflowOperation.Compare,
            InputPath = left,
            ComparisonPath = right,
            OutputPath = gallery,
            ConflictPolicy = OfficeWorkflowConflictPolicy.Fail
        });
        OfficeWorkflowResult sanitization = await runner.RunAsync(new OfficeWorkflowRequest {
            Operation = OfficeWorkflowOperation.Sanitize,
            InputPath = left,
            OutputPath = System.IO.Path.Combine(scope.Path, "sanitized.pdf"),
            ConflictPolicy = OfficeWorkflowConflictPolicy.Fail
        });

        Assert.True(comparison.Succeeded, comparison.Summary);
        Assert.True(comparison.HealthReport!.Verified);
        Assert.NotNull(comparison.HealthReport.After);
        Assert.True(File.Exists(gallery));
        Assert.True(sanitization.Succeeded, sanitization.Summary);
        Assert.NotNull(sanitization.HealthReport!.After);
        Assert.True(sanitization.HealthReport.Verified);
    }

    [Fact]
    public async Task ComparisonAcceptsIndependentPasswordsForEncryptedInputs() {
        using var scope = new TestDirectory();
        string left = System.IO.Path.Combine(scope.Path, "left-encrypted.pdf");
        string right = System.IO.Path.Combine(scope.Path, "right-encrypted.pdf");
        await File.WriteAllBytesAsync(
            left,
            PdfDocument.Create(
                compose => compose.Page(page => page.Content(content => content.Item(item => item.Paragraph(paragraph => paragraph.Text("Comparable content"))))),
                new PdfOptions().SetEncryption("left-open", "left-owner"))
                .ToBytes());
        await File.WriteAllBytesAsync(
            right,
            PdfDocument.Create(
                compose => compose.Page(page => page.Content(content => content.Item(item => item.Paragraph(paragraph => paragraph.Text("Comparable content"))))),
                new PdfOptions().SetEncryption("right-open", "right-owner"))
                .ToBytes());

        OfficeWorkflowResult result = await new OfficeWorkflowRunner().RunAsync(new OfficeWorkflowRequest {
            Operation = OfficeWorkflowOperation.Compare,
            InputPath = left,
            ComparisonPath = right,
            OutputPath = System.IO.Path.Combine(scope.Path, "encrypted-comparison.html"),
            ConflictPolicy = OfficeWorkflowConflictPolicy.Fail,
            PdfPassword = "left-open",
            ComparisonPdfPassword = "right-open"
        });

        Assert.True(result.Succeeded, result.Summary);
        Assert.True(result.HealthReport!.Before.CanRead);
        Assert.True(result.HealthReport.After!.CanRead);
        Assert.DoesNotContain(result.Diagnostics, diagnostic =>
            diagnostic.Message.Contains("left-open", StringComparison.Ordinal) ||
            diagnostic.Message.Contains("right-open", StringComparison.Ordinal));
    }

    private static OfficeWorkflowRequest Optimize(string input, string output, OfficeWorkflowConflictPolicy policy) => new() {
        Operation = OfficeWorkflowOperation.Optimize,
        InputPath = input,
        OutputPath = output,
        ConflictPolicy = policy
    };

    private sealed class UnderreportedLengthStream : Stream {
        private readonly MemoryStream _inner;
        private readonly long _reportedLength;

        internal UnderreportedLengthStream(byte[] bytes, long reportedLength) {
            _inner = new MemoryStream(bytes, writable: false);
            _reportedLength = reportedLength;
        }

        public override bool CanRead => true;
        public override bool CanSeek => true;
        public override bool CanWrite => false;
        public override long Length => _reportedLength;
        public override long Position { get => _inner.Position; set => _inner.Position = value; }
        public override void Flush() { }
        public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);
        public override long Seek(long offset, SeekOrigin origin) => _inner.Seek(offset, origin);
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        protected override void Dispose(bool disposing) {
            if (disposing) _inner.Dispose();
            base.Dispose(disposing);
        }
    }

    private static string CreateInput(string root, string routeId) {
        if (routeId.StartsWith("pdf-", StringComparison.Ordinal)) return CreatePdf(root, "source.pdf");
        switch (routeId) {
            case "docx-pdf": {
                string path = System.IO.Path.Combine(root, "source.docx");
                using WordDocument document = WordDocument.Create(path);
                document.AddParagraph("OfficeIMO workflow Word source");
                document.Save();
                return path;
            }
            case "xlsx-pdf": {
                string path = System.IO.Path.Combine(root, "source.xlsx");
                using ExcelDocument document = ExcelDocument.Create(path);
                document.AddWorksheet("Data").Cell(1, 1, "OfficeIMO workflow Excel source");
                document.Save();
                return path;
            }
            case "pptx-pdf": {
                string path = System.IO.Path.Combine(root, "source.pptx");
                using PowerPointPresentation presentation = PowerPointPresentation.Create(path);
                presentation.AddSlide().AddTextBoxPoints("OfficeIMO workflow PowerPoint source", 40, 40, 500, 60);
                presentation.Save();
                return path;
            }
            case "html-pdf": {
                string path = System.IO.Path.Combine(root, "source.html");
                File.WriteAllText(path, "<!doctype html><html><body><h1>OfficeIMO workflow HTML source</h1></body></html>", Encoding.UTF8);
                return path;
            }
            default:
                throw new ArgumentOutOfRangeException(nameof(routeId), routeId, "Unknown test route.");
        }
    }

    private static string CreatePdf(string root, string fileName, string text = "OfficeIMO workflow PDF source") {
        string path = System.IO.Path.Combine(root, fileName);
        PdfDocument.Create(compose => compose.Page(page => page.Content(content => content.Item(item => item.Paragraph(paragraph => paragraph.Text(text)))))).Save(path);
        return path;
    }

    private static string NormalizeExtension(string extension) => extension.StartsWith('.') ? extension : "." + extension;

    private sealed class InlineProgress<T>(Action<T> report) : IProgress<T> {
        public void Report(T value) => report(value);
    }

    private sealed class TestDirectory : IDisposable {
        public TestDirectory() {
            Path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "officeimo-workflows-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(Path);
        }

        public string Path { get; }

        public void Dispose() {
            try {
                Directory.Delete(Path, recursive: true);
            } catch (IOException) {
                // Test cleanup is best effort on Windows where package streams can briefly retain handles.
            } catch (UnauthorizedAccessException) {
                // Test cleanup is best effort on Windows where package streams can briefly retain handles.
            }
        }
    }
}
