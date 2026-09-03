using System.IO.Compression;
using System.Runtime.InteropServices;
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
        Assert.All(prepared, item => Assert.Equal(OfficeWorkflowConflictPolicy.Rename, item.ConflictPolicy));
        Assert.Same(prepared[0].BatchBlockedOutputIdentities, prepared[^1].BatchBlockedOutputIdentities);
        Assert.Equal(10_001, prepared[0].BatchBlockedOutputIdentities!.Count);
    }

    [Fact]
    public async Task BatchRenameRetriesWithoutClaimingAnotherRequestsReservation() {
        using var scope = new TempScope();
        string first = scope.Write("first.html", HtmlWithExternalManifest("first"));
        string second = scope.Write("second.html", HtmlWithExternalManifest("second"));
        string requested = Path.Combine(scope.Path, "cleaned.html");
        var progress = new CreatingOutputProgress("first", requested);

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
                OutputPath = requested,
                ConflictPolicy = OfficeWorkflowConflictPolicy.Rename
            }
        ], progress: progress);

        Assert.True(progress.Created);
        Assert.All(results, result => Assert.True(result.Succeeded, result.Summary));
        Assert.Equal(Path.Combine(scope.Path, "cleaned (2).html"), results[0].OutputPath);
        Assert.Equal(Path.Combine(scope.Path, "cleaned (1).html"), results[1].OutputPath);
        Assert.Equal("occupied during execution", File.ReadAllText(requested));
    }

    [Fact]
    public async Task BatchRejectsAncestorOutputPathsBeforePublishingAnything() {
        using var scope = new TempScope();
        string first = scope.Write("first.html", HtmlWithExternalManifest("first"));
        string second = scope.Write("second.html", HtmlWithExternalManifest("second"));
        string parentOutput = Path.Combine(scope.Path, "result.html");
        string childOutput = Path.Combine(parentOutput, "child.html");

        ArgumentException exception = await Assert.ThrowsAsync<ArgumentException>(() =>
            new OfficeWorkflowRunner().RunProvenanceBatchAsync([
                new OfficeProvenanceWorkflowRequest {
                    Id = "first",
                    Operation = OfficeProvenanceWorkflowOperation.Remove,
                    InputPath = first,
                    OutputPath = parentOutput,
                    ConflictPolicy = OfficeWorkflowConflictPolicy.Replace
                },
                new OfficeProvenanceWorkflowRequest {
                    Id = "second",
                    Operation = OfficeProvenanceWorkflowOperation.Remove,
                    InputPath = second,
                    OutputPath = childOutput,
                    ConflictPolicy = OfficeWorkflowConflictPolicy.Replace
                }
            ]));

        Assert.Contains("ancestor/descendant", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.False(File.Exists(parentOutput));
        Assert.False(Directory.Exists(parentOutput));
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

    [Fact]
    public async Task AssessmentReportsTheLogicalSourcePathWhileReadingTheSnapshot() {
        using var scope = new TempScope();
        string input = scope.Write("page.html", "<!doctype html><html><body>review\u200Bthis</body></html>");

        OfficeProvenanceWorkflowResult result = await new OfficeWorkflowRunner().RunProvenanceAsync(
            new OfficeProvenanceWorkflowRequest {
                Operation = OfficeProvenanceWorkflowOperation.Assess,
                InputPath = input
            });

        Assert.True(result.Succeeded, result.Summary);
        OfficeTextIntegrityFinding finding = Assert.Single(result.Assessment!.TextIntegrity!.Findings);
        Assert.Equal(Path.GetFullPath(input), finding.Location);
        Assert.DoesNotContain("officeimo-provenance-", finding.Location, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task InspectReentersTheDetectedHtmlOwnerForAnUnknownExtension() {
        using var scope = new TempScope();
        string input = scope.Write("asset.bin", HtmlWithExternalManifest("body"));

        OfficeProvenanceWorkflowResult result = await new OfficeWorkflowRunner().RunProvenanceAsync(
            new OfficeProvenanceWorkflowRequest {
                Operation = OfficeProvenanceWorkflowOperation.Inspect,
                InputPath = input
            });

        Assert.True(result.Succeeded, result.Summary);
        Assert.Equal("OfficeIMO.Html", result.OwnerPackage);
        Assert.Equal(OfficeProvenanceAssetFormat.Html, result.Inspection!.Format);
        Assert.Equal(OfficeProvenanceCarrierKind.C2paExternalManifest, Assert.Single(result.Inspection.Evidence).Carrier);
    }

    [Fact]
    public async Task AssessmentWithoutProvidersDoesNotCopyAnOversizedExternalManifest() {
        using var scope = new TempScope();
        string input = scope.Write("page.html", HtmlWithExternalManifest("body"));
        scope.Write("claim.c2pa", new string('x', 128));
        var assessment = new OfficeProvenanceAssessmentOptions();
        assessment.Structural.MaxManifestBytes = 8;

        OfficeProvenanceWorkflowResult result = await new OfficeWorkflowRunner().RunProvenanceAsync(
            new OfficeProvenanceWorkflowRequest {
                Operation = OfficeProvenanceWorkflowOperation.Assess,
                InputPath = input,
                Assessment = assessment
            });

        Assert.True(result.Succeeded, result.Summary);
        Assert.Single(result.Assessment!.Structural.Evidence);
        Assert.Null(result.Assessment.Verification);
        Assert.Empty(result.Assessment.ProviderSignals);
    }

    [Fact]
    public async Task AssessmentSnapshotsRelativeExternalManifestDependenciesForProviders() {
        using var scope = new TempScope();
        string input = scope.Write("page.html", HtmlWithExternalManifest("body"));
        scope.Write("claim.c2pa", "immutable claim");
        var verifier = new RelativeManifestVerifier();

        OfficeProvenanceWorkflowResult result = await new OfficeWorkflowRunner(verifier).RunProvenanceAsync(
            new OfficeProvenanceWorkflowRequest {
                Operation = OfficeProvenanceWorkflowOperation.Assess,
                InputPath = input
            });

        Assert.True(result.Succeeded, result.Summary);
        Assert.Equal(OfficeProvenanceVerificationStatus.Valid, result.Assessment!.Verification!.Status);
        Assert.True(verifier.SawRelativeManifest);
        Assert.NotEqual(Path.GetDirectoryName(input), verifier.ObservedDirectory);
        Assert.False(Directory.Exists(verifier.ObservedDirectory));
    }

    [Fact]
    public async Task AssessmentFailsClosedWhenAProviderReplacesACapturedExternalManifest() {
        using var scope = new TempScope();
        string input = scope.Write("page.html", HtmlWithExternalManifest("body"));
        scope.Write("claim.c2pa", "immutable claim");
        var verifier = new ReplacingRelativeManifestVerifier();

        OfficeProvenanceWorkflowResult result = await new OfficeWorkflowRunner(verifier).RunProvenanceAsync(
            new OfficeProvenanceWorkflowRequest {
                Operation = OfficeProvenanceWorkflowOperation.Assess,
                InputPath = input
            });

        Assert.True(verifier.Replaced || verifier.MutationBlocked);
        if (verifier.Replaced) {
            Assert.False(result.Succeeded);
            Assert.Contains("external provenance manifest changed", result.Summary, StringComparison.OrdinalIgnoreCase);
        } else {
            Assert.True(result.Succeeded, result.Summary);
        }
        Assert.False(Directory.Exists(verifier.ObservedDirectory));
    }

    [Fact]
    public async Task AssessmentResolvesRelativeManifestAgainstTheHtmlBaseElement() {
        using var scope = new TempScope();
        string input = scope.Write(
            "page.html",
            "<!doctype html><html><head><base href=\"sub/\"><link rel=\"c2pa-manifest\" href=\"claim.c2pa\"></head><body>body</body></html>");
        Directory.CreateDirectory(Path.Combine(scope.Path, "sub"));
        File.WriteAllText(Path.Combine(scope.Path, "sub", "claim.c2pa"), "base-relative claim");
        var verifier = new BaseRelativeManifestVerifier();

        OfficeProvenanceWorkflowResult result = await new OfficeWorkflowRunner(verifier).RunProvenanceAsync(
            new OfficeProvenanceWorkflowRequest {
                Operation = OfficeProvenanceWorkflowOperation.Assess,
                InputPath = input
            });

        Assert.True(result.Succeeded, result.Summary);
        Assert.True(verifier.SawRelativeManifest);
        Assert.False(Directory.Exists(verifier.ObservedDirectory));
    }

    [Fact]
    public async Task AssessmentFailsClosedWhenAProviderReplacesThePrimarySnapshot() {
#if NET8_0_OR_GREATER
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return;
        using var scope = new TempScope();
        string input = scope.Write("page.html", HtmlWithExternalManifest("body"));
        var detector = new ReplacingSnapshotDetector();

        OfficeProvenanceWorkflowResult result = await new OfficeWorkflowRunner(
            provenanceVerifier: null,
            provenanceSignalDetectors: [detector]).RunProvenanceAsync(
            new OfficeProvenanceWorkflowRequest {
                Operation = OfficeProvenanceWorkflowOperation.Assess,
                InputPath = input
            });

        Assert.True(detector.Replaced);
        Assert.False(result.Succeeded);
        Assert.Contains("primary provenance snapshot changed", result.Summary, StringComparison.OrdinalIgnoreCase);
#endif
    }

    [Fact]
    public async Task AssessmentSharesExpandedDataBudgetWithExternalManifestCapture() {
        using var scope = new TempScope();
        string input = scope.Write(
            "page.html",
            "<!doctype html><html><head><link rel=\"c2pa-manifest\" href=\"claim.c2pa\"></head>" +
            "<body><img src=\"data:image/png;base64,AQIDBA==\"></body></html>");
        scope.Write("claim.c2pa", "four");
        var assessment = new OfficeProvenanceAssessmentOptions();
        assessment.Structural.MaxExpandedContainerBytes = 7;
        var verifier = new RelativeManifestVerifier();

        OfficeProvenanceWorkflowResult result = await new OfficeWorkflowRunner(verifier).RunProvenanceAsync(
            new OfficeProvenanceWorkflowRequest {
                Operation = OfficeProvenanceWorkflowOperation.Assess,
                InputPath = input,
                Assessment = assessment
            });

        Assert.False(result.Succeeded);
        Assert.Contains("expanded-data limit", result.Summary, StringComparison.OrdinalIgnoreCase);
        Assert.False(verifier.SawRelativeManifest);
    }

    [Fact]
    public async Task DiskBackedSnapshotAcceptsWorkflowCeilingsAboveInt32() {
        using var scope = new TempScope();
        string input = scope.Write("page.html", HtmlWithExternalManifest("body"));

        OfficeProvenanceWorkflowResult result = await new OfficeWorkflowRunner().RunProvenanceAsync(
            new OfficeProvenanceWorkflowRequest {
                Operation = OfficeProvenanceWorkflowOperation.Assess,
                InputPath = input,
                Limits = new OfficeWorkflowLimits {
                    MaximumInputBytes = (long)int.MaxValue + 4096L,
                    MaximumOutputBytes = (long)int.MaxValue + 4096L
                }
            });

        Assert.True(result.Succeeded, result.Summary);
    }

    [Fact]
    public async Task PublicationRejectsAStagedArtifactChangedAfterValidation() {
        using var scope = new TempScope();
        string input = scope.Write("page.html", HtmlWithExternalManifest("original"));
        string output = Path.Combine(scope.Path, "cleaned.html");
        var progress = new ReplacingStagedArtifactProgress(scope.Path, output);

        OfficeProvenanceWorkflowResult result = await new OfficeWorkflowRunner().RunProvenanceAsync(
            new OfficeProvenanceWorkflowRequest {
                Operation = OfficeProvenanceWorkflowOperation.Remove,
                InputPath = input,
                OutputPath = output
            },
            progress);

        Assert.False(result.Succeeded);
        Assert.Equal(OfficeWorkflowFailureKind.OutputFailed, result.FailureKind);
        Assert.Contains("changed after output validation", result.Summary, StringComparison.OrdinalIgnoreCase);
        Assert.True(progress.Replaced);
        Assert.False(File.Exists(output));
    }

    [Fact]
    public async Task ReplaceRestoresTheDisplacedDestinationWhenFinalValidationIsCancelled() {
        using var scope = new TempScope();
        string input = scope.Write("page.html", HtmlWithExternalManifest("original"));
        string output = scope.Write("cleaned.html", "existing destination");
        using var cancellation = new CancellationTokenSource();
        var progress = new CancellingProgress("replace-finalization", cancellation);

        OfficeProvenanceWorkflowResult result = await new OfficeWorkflowRunner().RunProvenanceAsync(
            new OfficeProvenanceWorkflowRequest {
                Id = "replace-finalization",
                Operation = OfficeProvenanceWorkflowOperation.Remove,
                InputPath = input,
                OutputPath = output,
                ConflictPolicy = OfficeWorkflowConflictPolicy.Replace
            },
            progress,
            cancellation.Token);

        Assert.Equal(OfficeWorkflowStatus.Cancelled, result.Status);
        Assert.True(cancellation.IsCancellationRequested);
        Assert.Equal("existing destination", File.ReadAllText(output));
    }

    [Fact]
    public async Task ReplaceRestoresTheDisplacedDestinationWhenPublishedArtifactDisappears() {
        using var scope = new TempScope();
        string input = scope.Write("page.html", HtmlWithExternalManifest("original"));
        string output = scope.Write("cleaned.html", "existing destination");
        var progress = new DeletingPublishedArtifactProgress(output);

        OfficeProvenanceWorkflowResult result = await new OfficeWorkflowRunner().RunProvenanceAsync(
            new OfficeProvenanceWorkflowRequest {
                Operation = OfficeProvenanceWorkflowOperation.Remove,
                InputPath = input,
                OutputPath = output,
                ConflictPolicy = OfficeWorkflowConflictPolicy.Replace
            },
            progress);

        Assert.False(result.Succeeded);
        Assert.True(progress.Deleted);
        Assert.Equal("existing destination", File.ReadAllText(output));
    }

    [Fact]
    public async Task ReplaceRejectsAnExistingDestinationAboveTheOutputLimitWithoutHashingOrChangingIt() {
        using var scope = new TempScope();
        string input = scope.Write("page.html", HtmlWithExternalManifest("original"));
        string existing = new string('x', 2048);
        string output = scope.Write("cleaned.html", existing);

        OfficeProvenanceWorkflowResult result = await new OfficeWorkflowRunner().RunProvenanceAsync(
            new OfficeProvenanceWorkflowRequest {
                Operation = OfficeProvenanceWorkflowOperation.Remove,
                InputPath = input,
                OutputPath = output,
                ConflictPolicy = OfficeWorkflowConflictPolicy.Replace,
                Limits = new OfficeWorkflowLimits {
                    MaximumInputBytes = 4096,
                    MaximumOutputBytes = 1024
                }
            });

        Assert.False(result.Succeeded);
        Assert.Equal(OfficeWorkflowFailureKind.OutputFailed, result.FailureKind);
        Assert.Contains("existing provenance destination", result.Summary, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(existing, File.ReadAllText(output));
    }

    [Theory]
    [InlineData(typeof(IOException))]
    [InlineData(typeof(UnauthorizedAccessException))]
    public void InputAccessFailuresUseTheInputFailureContract(Type exceptionType) {
        var exception = (Exception)Activator.CreateInstance(exceptionType, "input unavailable")!;

        OfficeWorkflowFailureKind kind = OfficeWorkflowRunner.ClassifyFailure(
            exception,
            OfficeWorkflowRunner.WorkflowFailureStage.Input);

        Assert.Equal(OfficeWorkflowFailureKind.UnsupportedInput, kind);
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

    private sealed class RelativeManifestVerifier : IOfficeProvenanceVerifier {
        public string Name => "relative-manifest";
        internal bool SawRelativeManifest { get; private set; }
        internal string? ObservedDirectory { get; private set; }

        public OfficeProvenanceVerificationResult Verify(
            string filePath,
            OfficeProvenanceVerificationOptions? options = null) {
            ObservedDirectory = Path.GetDirectoryName(Path.GetFullPath(filePath));
            string manifestPath = Path.Combine(ObservedDirectory!, "claim.c2pa");
            SawRelativeManifest = File.Exists(manifestPath) && File.ReadAllText(manifestPath) == "immutable claim";
            return new OfficeProvenanceVerificationResult(
                SawRelativeManifest ? OfficeProvenanceVerificationStatus.Valid : OfficeProvenanceVerificationStatus.Invalid,
                Name,
                Array.Empty<string>());
        }
    }

    private sealed class ReplacingRelativeManifestVerifier : IOfficeProvenanceVerifier {
        public string Name => "replacing-relative-manifest";
        internal bool Replaced { get; private set; }
        internal bool MutationBlocked { get; private set; }
        internal string? ObservedDirectory { get; private set; }

        public OfficeProvenanceVerificationResult Verify(
            string filePath,
            OfficeProvenanceVerificationOptions? options = null) {
            ObservedDirectory = Path.GetDirectoryName(Path.GetFullPath(filePath));
            string manifestPath = Path.Combine(ObservedDirectory!, "claim.c2pa");
            string replacementPath = manifestPath + ".replacement";
            try {
                File.WriteAllText(replacementPath, "replaced claim");
                File.Move(replacementPath, manifestPath, overwrite: true);
                Replaced = true;
            } catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) {
                MutationBlocked = true;
                if (File.Exists(replacementPath)) File.Delete(replacementPath);
            }
            return new OfficeProvenanceVerificationResult(
                OfficeProvenanceVerificationStatus.Valid,
                Name,
                Array.Empty<string>());
        }
    }

    private sealed class BaseRelativeManifestVerifier : IOfficeProvenanceVerifier {
        public string Name => "base-relative-manifest";
        internal bool SawRelativeManifest { get; private set; }
        internal string? ObservedDirectory { get; private set; }

        public OfficeProvenanceVerificationResult Verify(
            string filePath,
            OfficeProvenanceVerificationOptions? options = null) {
            ObservedDirectory = Path.GetDirectoryName(Path.GetFullPath(filePath));
            string manifestPath = Path.Combine(ObservedDirectory!, "sub", "claim.c2pa");
            SawRelativeManifest = File.Exists(manifestPath) &&
                                  File.ReadAllText(manifestPath) == "base-relative claim";
            return new OfficeProvenanceVerificationResult(
                SawRelativeManifest ? OfficeProvenanceVerificationStatus.Valid : OfficeProvenanceVerificationStatus.Invalid,
                Name,
                Array.Empty<string>());
        }
    }

    private sealed class CreatingOutputProgress(string requestId, string outputPath) : IProgress<OfficeWorkflowProgress> {
        internal bool Created { get; private set; }

        public void Report(OfficeWorkflowProgress value) {
            if (Created || value.RequestId != requestId || value.Stage != "validate") return;
            File.WriteAllText(outputPath, "occupied during execution");
            Created = true;
        }
    }

#if NET8_0_OR_GREATER
    private sealed class ReplacingSnapshotDetector : IOfficeProvenanceSignalDetector {
        public string Name => "replacing-snapshot";
        public OfficeProvenanceSignalKind SignalKind => OfficeProvenanceSignalKind.DeterministicArtifact;
        internal bool Replaced { get; private set; }

        public OfficeProvenanceSignalResult Detect(string filePath) {
            string replacementPath = filePath + ".replacement";
            File.WriteAllText(replacementPath, "replacement");
            File.Move(replacementPath, filePath, overwrite: true);
            Replaced = true;
            return new OfficeProvenanceSignalResult(
                Name,
                SignalKind,
                OfficeProvenanceSignalStatus.Detected);
        }
    }
#endif

    private sealed class ReplacingStagedArtifactProgress : IProgress<OfficeWorkflowProgress> {
        private readonly string _directory;
        private readonly string _outputPath;

        internal ReplacingStagedArtifactProgress(string directory, string outputPath) {
            _directory = directory;
            _outputPath = outputPath;
        }

        internal bool Replaced { get; private set; }

        public void Report(OfficeWorkflowProgress value) {
            if (Replaced || !string.Equals(value.Stage, "publish", StringComparison.Ordinal)) return;
            string stagingPath = Directory.GetFiles(
                    _directory,
                    "." + Path.GetFileNameWithoutExtension(_outputPath) + ".*" + Path.GetExtension(_outputPath))
                .Single();
            File.WriteAllText(stagingPath, "<!doctype html><html><body>replacement</body></html>");
            Replaced = true;
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

    private sealed class DeletingPublishedArtifactProgress : IProgress<OfficeWorkflowProgress> {
        private readonly string _path;

        internal DeletingPublishedArtifactProgress(string path) => _path = path;

        internal bool Deleted { get; private set; }

        public void Report(OfficeWorkflowProgress value) {
            if (Deleted || !string.Equals(value.Stage, "complete", StringComparison.Ordinal)) return;
            File.Delete(_path);
            Deleted = true;
        }
    }

}
