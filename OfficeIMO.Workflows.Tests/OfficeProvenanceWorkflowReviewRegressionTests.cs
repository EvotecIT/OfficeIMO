using System.IO.Compression;
using System.Text;
using OfficeIMO.OpenDocument;
using OfficeIMO.Provenance;

namespace OfficeIMO.Workflows.Tests;

public sealed partial class OfficeProvenanceWorkflowTests {
    [Fact]
    public async Task RemovalCanReopenAnExpandedOutputWithinItsSeparateOutputBudget() {
        using var scope = new TempScope();
        string input = scope.Write("compact.html", "<link rel=c2pa-manifest href=x>");
        string probeOutput = Path.Combine(scope.Path, "probe.html");
        OfficeProvenanceWorkflowResult probe = await new OfficeWorkflowRunner().RunProvenanceAsync(
            new OfficeProvenanceWorkflowRequest {
                Operation = OfficeProvenanceWorkflowOperation.Remove,
                InputPath = input,
                OutputPath = probeOutput
            });
        Assert.True(probe.Succeeded, probe.Summary);
        long inputBytes = new FileInfo(input).Length;
        long outputBytes = new FileInfo(probeOutput).Length;
        Assert.True(outputBytes > inputBytes, $"Expected HTML normalization to expand {inputBytes} bytes, but received {outputBytes} bytes.");

        string output = Path.Combine(scope.Path, "cleaned.html");
        OfficeProvenanceWorkflowResult result = await new OfficeWorkflowRunner().RunProvenanceAsync(
            new OfficeProvenanceWorkflowRequest {
                Operation = OfficeProvenanceWorkflowOperation.Remove,
                InputPath = input,
                OutputPath = output,
                Limits = new OfficeWorkflowLimits {
                    MaximumInputBytes = inputBytes,
                    MaximumOutputBytes = outputBytes
                }
            });

        Assert.True(result.Succeeded, result.Summary);
        Assert.Equal(outputBytes, result.OutputBytes);
    }

    [Fact]
    public async Task RemovalRejectsAnUnchangedArtifactAboveItsIndependentOutputBudget() {
        using var scope = new TempScope();
        string input = scope.Write("unchanged.html", "<!doctype html><html><body>unchanged</body></html>");
        string output = Path.Combine(scope.Path, "cleaned.html");
        long inputBytes = new FileInfo(input).Length;

        OfficeProvenanceWorkflowResult result = await new OfficeWorkflowRunner().RunProvenanceAsync(
            new OfficeProvenanceWorkflowRequest {
                Operation = OfficeProvenanceWorkflowOperation.Remove,
                InputPath = input,
                OutputPath = output,
                Limits = new OfficeWorkflowLimits {
                    MaximumInputBytes = inputBytes,
                    MaximumOutputBytes = inputBytes - 1L
                }
            });

        Assert.False(result.Succeeded);
        Assert.Contains("output limit", result.Summary, StringComparison.OrdinalIgnoreCase);
        Assert.False(File.Exists(output));
    }

    [Fact]
    public async Task RemovalUsesThePreflightSnapshotWhenTheSourceIsReplaced() {
        using var scope = new TempScope();
        string input = scope.Write("source.html", HtmlWithExternalManifest("original"));
        string output = Path.Combine(scope.Path, "cleaned.html");
        long inputBytes = new FileInfo(input).Length;
        var progress = new ReplacingRemovalProgress(
            input,
            "<!doctype html><html><body>replacement" + new string('x', 4096) + "</body></html>");

        OfficeProvenanceWorkflowResult result = await new OfficeWorkflowRunner().RunProvenanceAsync(
            new OfficeProvenanceWorkflowRequest {
                Operation = OfficeProvenanceWorkflowOperation.Remove,
                InputPath = input,
                OutputPath = output,
                Limits = new OfficeWorkflowLimits {
                    MaximumInputBytes = inputBytes,
                    MaximumOutputBytes = 16 * 1024
                }
            },
            progress);

        Assert.True(result.Succeeded, result.Summary);
        string cleaned = File.ReadAllText(output);
        Assert.Contains("original", cleaned, StringComparison.Ordinal);
        Assert.DoesNotContain("replacement", cleaned, StringComparison.Ordinal);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "RemovalSnapshot");
    }

    [Theory]
    [InlineData(".odt", OdfMediaTypes.Text)]
    [InlineData(".ods", OdfMediaTypes.Spreadsheet)]
    [InlineData(".odp", OdfMediaTypes.Presentation)]
    [InlineData(".odg", OdfMediaTypes.Graphics)]
    [InlineData(".ott", OdfMediaTypes.TextTemplate)]
    [InlineData(".ots", OdfMediaTypes.SpreadsheetTemplate)]
    [InlineData(".otp", OdfMediaTypes.PresentationTemplate)]
    [InlineData(".otg", OdfMediaTypes.GraphicsTemplate)]
    public async Task EveryAdvertisedOpenDocumentExtensionCanBeInspectedAndRemoved(
        string extension,
        string mediaType) {
        using var scope = new TempScope();
        string input = Path.Combine(scope.Path, "document" + extension);
        string output = Path.Combine(scope.Path, "cleaned" + extension);
        CreateOpenDocumentPackage(input, mediaType);

        OfficeProvenanceWorkflowResult inspection = await new OfficeWorkflowRunner().RunProvenanceAsync(
            new OfficeProvenanceWorkflowRequest {
                Operation = OfficeProvenanceWorkflowOperation.Inspect,
                InputPath = input
            });
        OfficeProvenanceWorkflowResult removal = await new OfficeWorkflowRunner().RunProvenanceAsync(
            new OfficeProvenanceWorkflowRequest {
                Operation = OfficeProvenanceWorkflowOperation.Remove,
                InputPath = input,
                OutputPath = output
            });

        Assert.True(inspection.Succeeded, inspection.Summary);
        Assert.Equal("OfficeIMO.OpenDocument", inspection.OwnerPackage);
        Assert.True(removal.Succeeded, removal.Summary);
        Assert.True(File.Exists(output));
    }

    [Fact]
    public async Task BatchReservesRenameDestinationBeforeLaterReplacementRuns() {
        using var scope = new TempScope();
        string first = scope.Write("first.html", HtmlWithExternalManifest("first"));
        string second = scope.Write("second.html", HtmlWithExternalManifest("second"));
        string requested = scope.Write("clean.html", "occupied");
        string laterOutput = Path.Combine(scope.Path, "clean (1).html");

        IReadOnlyList<OfficeProvenanceWorkflowResult> results = await new OfficeWorkflowRunner().RunProvenanceBatchAsync([
            new OfficeProvenanceWorkflowRequest {
                Id = "first",
                Operation = OfficeProvenanceWorkflowOperation.Remove,
                InputPath = first,
                OutputPath = requested,
                ConflictPolicy = OfficeWorkflowConflictPolicy.Rename
            },
            new OfficeProvenanceWorkflowRequest {
                Id = "second",
                Operation = OfficeProvenanceWorkflowOperation.Remove,
                InputPath = second,
                OutputPath = laterOutput,
                ConflictPolicy = OfficeWorkflowConflictPolicy.Replace
            }
        ]);

        Assert.Collection(
            results,
            result => Assert.Equal(Path.Combine(scope.Path, "clean (2).html"), result.OutputPath),
            result => Assert.Equal(laterOutput, result.OutputPath));
        Assert.Contains("first", File.ReadAllText(results[0].OutputPath!), StringComparison.Ordinal);
        Assert.Contains("second", File.ReadAllText(results[1].OutputPath!), StringComparison.Ordinal);
        Assert.Equal("occupied", File.ReadAllText(requested));
    }

    [Fact]
    public async Task BatchKeepsSingleRequestInPlaceReplacementAvailable() {
        using var scope = new TempScope();
        string input = scope.Write("page.html", HtmlWithExternalManifest("body"));

        IReadOnlyList<OfficeProvenanceWorkflowResult> results = await new OfficeWorkflowRunner().RunProvenanceBatchAsync([
            new OfficeProvenanceWorkflowRequest {
                Operation = OfficeProvenanceWorkflowOperation.Remove,
                InputPath = input,
                OutputPath = input,
                ConflictPolicy = OfficeWorkflowConflictPolicy.Replace
            }
        ]);

        OfficeProvenanceWorkflowResult result = Assert.Single(results);
        Assert.True(result.Succeeded, result.Summary);
        Assert.Equal(input, result.OutputPath);
        Assert.DoesNotContain("c2pa-manifest", File.ReadAllText(input), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void BatchRenamePlanningAdvancesOneDestinationFamilyLinearly() {
        using var scope = new TempScope();
        string input = scope.Write("input.html", HtmlWithExternalManifest("body"));
        string requested = Path.Combine(scope.Path, "cleaned.html");
        OfficeProvenanceWorkflowRequest[] requests = Enumerable.Range(0, 10_000)
            .Select(index => new OfficeProvenanceWorkflowRequest {
                Id = "request-" + index,
                Operation = OfficeProvenanceWorkflowOperation.Remove,
                InputPath = input,
                OutputPath = requested,
                ConflictPolicy = OfficeWorkflowConflictPolicy.Rename
            })
            .ToArray();

        OfficeProvenanceWorkflowRequest[] prepared = OfficeWorkflowRunner.PrepareBatchRemovalPaths(requests);

        Assert.Equal(requested, prepared[0].OutputPath);
        Assert.Equal(Path.Combine(scope.Path, "cleaned (9999).html"), prepared[^1].OutputPath);
        Assert.Equal(10_000, prepared.Select(item => item.OutputPath).Distinct(StringComparer.OrdinalIgnoreCase).Count());
        Assert.All(prepared, item => Assert.Equal(OfficeWorkflowConflictPolicy.Fail, item.ConflictPolicy));
    }

    [Fact]
    public void BatchRenamePlanningHonorsPreCancelledWorkWithoutFilesystemScanning() {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        var request = new OfficeProvenanceWorkflowRequest {
            Operation = OfficeProvenanceWorkflowOperation.Remove,
            InputPath = "not-inspected-input.html",
            OutputPath = "not-inspected-output.html",
            ConflictPolicy = OfficeWorkflowConflictPolicy.Rename
        };

        OfficeProvenanceWorkflowRequest[] prepared = OfficeWorkflowRunner.PrepareBatchRemovalPaths(
            [request],
            cancellation.Token);

        Assert.Same(request, Assert.Single(prepared));
    }

    [Fact]
    public void PdfOutputLimitMarkersUseTheWorkflowOutputFailureContract() {
        InvalidDataException failure = OfficeIMO.Pdf.PdfOutputLimitErrors.Create(
            "The rewritten PDF exceeds the configured output limit.");

        OfficeWorkflowFailureKind kind = OfficeWorkflowRunner.ClassifyFailure(
            failure,
            OfficeWorkflowRunner.WorkflowFailureStage.Operation);

        Assert.Equal(OfficeWorkflowFailureKind.OutputFailed, kind);
    }

    [Fact]
    public async Task AssessmentComposesEverySignalFromOneImmutableSnapshot() {
        using var scope = new TempScope();
        string input = scope.Write("page.html", HtmlWithExternalManifest("original"));
        var detector = new ReplacingInputDetector(input);

        OfficeProvenanceWorkflowResult result = await new OfficeWorkflowRunner(
            provenanceVerifier: null,
            provenanceSignalDetectors: [detector]).RunProvenanceAsync(
            new OfficeProvenanceWorkflowRequest {
                Operation = OfficeProvenanceWorkflowOperation.Assess,
                InputPath = input
            });

        Assert.True(result.Succeeded, result.Summary);
        Assert.Single(result.Assessment!.Structural.Evidence);
        Assert.Equal(OfficeProvenanceSignalStatus.Detected, Assert.Single(result.Assessment.ProviderSignals).Status);
        Assert.NotNull(detector.ObservedPath);
        Assert.NotEqual(Path.GetFullPath(input), detector.ObservedPath);
        Assert.False(File.Exists(detector.ObservedPath));
        Assert.DoesNotContain("c2pa-manifest", File.ReadAllText(input), StringComparison.OrdinalIgnoreCase);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "AssessmentSnapshot");
    }

    private static void CreateOpenDocumentPackage(string path, string mediaType) {
        using var stream = File.Create(path);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: false);
        WriteEntry(archive, "mimetype", mediaType, CompressionLevel.NoCompression);
        WriteEntry(
            archive,
            "content.xml",
            "<?xml version=\"1.0\"?><office:document-content xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\"><office:body/></office:document-content>");
        WriteEntry(
            archive,
            "META-INF/manifest.xml",
            "<?xml version=\"1.0\"?><manifest:manifest xmlns:manifest=\"urn:oasis:names:tc:opendocument:xmlns:manifest:1.0\">" +
            "<manifest:file-entry manifest:full-path=\"/\" manifest:media-type=\"" + mediaType + "\"/>" +
            "<manifest:file-entry manifest:full-path=\"content.xml\" manifest:media-type=\"text/xml\"/>" +
            "</manifest:manifest>");
    }

    private sealed class ReplacingInputDetector : IOfficeProvenanceSignalDetector {
        private readonly string _originalPath;

        internal ReplacingInputDetector(string originalPath) => _originalPath = originalPath;

        public string Name => "replacing-input";
        public OfficeProvenanceSignalKind SignalKind => OfficeProvenanceSignalKind.DeterministicArtifact;
        internal string? ObservedPath { get; private set; }

        public OfficeProvenanceSignalResult Detect(string filePath) {
            ObservedPath = Path.GetFullPath(filePath);
            File.WriteAllText(_originalPath, "<!doctype html><html><body>replacement</body></html>");
            bool detected = File.ReadAllText(filePath).Contains("c2pa-manifest", StringComparison.OrdinalIgnoreCase);
            return new OfficeProvenanceSignalResult(
                Name,
                SignalKind,
                detected ? OfficeProvenanceSignalStatus.Detected : OfficeProvenanceSignalStatus.NotDetected);
        }
    }

    private sealed class ReplacingRemovalProgress : IProgress<OfficeWorkflowProgress> {
        private readonly string _path;
        private readonly string _replacement;
        private bool _replaced;

        internal ReplacingRemovalProgress(string path, string replacement) {
            _path = path;
            _replacement = replacement;
        }

        public void Report(OfficeWorkflowProgress value) {
            if (_replaced || !string.Equals(value.Stage, "remove", StringComparison.Ordinal)) return;
            File.WriteAllText(_path, _replacement);
            _replaced = true;
        }
    }
}
