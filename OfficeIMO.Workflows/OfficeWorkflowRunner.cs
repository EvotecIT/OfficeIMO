using System.Diagnostics;
using System.Text;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

namespace OfficeIMO.Workflows;

/// <summary>
/// Runs bounded local conversion and PDF health workflows and publishes validated artifacts atomically.
/// </summary>
public sealed partial class OfficeWorkflowRunner : IOfficeWorkflowRunner {
    /// <summary>Runs one workflow request.</summary>
    public async Task<OfficeWorkflowResult> RunAsync(
        OfficeWorkflowRequest request,
        IProgress<OfficeWorkflowProgress>? progress = null,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(request);
        var stopwatch = Stopwatch.StartNew();
        var diagnostics = new List<OfficeWorkflowDiagnostic>();
        long inputBytes = 0;
        string? stagingPath = null;

        try {
            ValidatedRequest validated = ValidateRequest(request);
            Report(progress, request.Id, "validate", "Validating input and workflow limits", 0.05D);
            cancellationToken.ThrowIfCancellationRequested();
            inputBytes = new FileInfo(validated.InputPath).Length;
            EnforceInputLimit(validated.InputPath, inputBytes, validated.Limits);
            if (validated.ComparisonPath is not null) {
                EnforceInputLimit(validated.ComparisonPath, new FileInfo(validated.ComparisonPath).Length, validated.Limits);
            }

            Report(progress, request.Id, "execute", DescribeOperation(validated.Operation), 0.18D);
            OperationArtifact artifact = await Task.Run(
                () => Execute(validated, diagnostics, cancellationToken),
                cancellationToken).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();

            if (artifact.Bytes is null) {
                Report(progress, request.Id, "complete", "Workflow report is ready", 1D);
                return CreateResult(
                    validated,
                    OfficeWorkflowStatus.Completed,
                    outputPath: null,
                    inputBytes,
                    outputBytes: 0,
                    stopwatch.Elapsed,
                    artifact.Summary,
                    diagnostics,
                    artifact.HealthReport);
            }

            if (artifact.Bytes.LongLength > validated.Limits.MaximumOutputBytes) {
                throw new InvalidOperationException(
                    $"Generated artifact is {artifact.Bytes.LongLength:N0} bytes, above the configured {validated.Limits.MaximumOutputBytes:N0}-byte limit.");
            }

            EnsureVerifiedHealthArtifact(validated.Operation, artifact.HealthReport);

            cancellationToken.ThrowIfCancellationRequested();
            string outputDirectory = Path.GetDirectoryName(validated.OutputPath!)!;
            Directory.CreateDirectory(outputDirectory);
            stagingPath = Path.Combine(
                outputDirectory,
                "." + Path.GetFileName(validated.OutputPath) + "." + Guid.NewGuid().ToString("N") + ".tmp");
            await File.WriteAllBytesAsync(stagingPath, artifact.Bytes, cancellationToken).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();

            Report(progress, request.Id, "validate-output", "Reopening the staged artifact", 0.72D);
            ValidateStagedArtifact(stagingPath, validated.OutputPath!, validated.PdfReadOptions);
            diagnostics.Add(new OfficeWorkflowDiagnostic(
                "OutputReopened",
                "The staged file was reopened successfully through its first-party OfficeIMO document API.",
                stage: "validate-output",
                details: new Dictionary<string, string>(StringComparer.Ordinal) {
                    ["stagedBytes"] = new FileInfo(stagingPath).Length.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    ["format"] = Path.GetExtension(validated.OutputPath!).ToLowerInvariant()
                }));
            cancellationToken.ThrowIfCancellationRequested();

            Report(progress, request.Id, "publish", "Publishing the validated artifact", 0.9D);
            string publishedPath = Publish(stagingPath, validated.OutputPath!, validated.ConflictPolicy);
            stagingPath = null;
            long outputBytes = new FileInfo(publishedPath).Length;
            diagnostics.Add(new OfficeWorkflowDiagnostic(
                "AtomicPublication",
                "The artifact was staged in the destination directory and published with one filesystem move.",
                stage: "publish"));
            Report(progress, request.Id, "complete", "Workflow completed", 1D);
            return CreateResult(
                validated,
                OfficeWorkflowStatus.Completed,
                publishedPath,
                inputBytes,
                outputBytes,
                stopwatch.Elapsed,
                artifact.Summary,
                diagnostics,
                artifact.HealthReport);
        } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
            diagnostics.Add(new OfficeWorkflowDiagnostic(
                "Cancelled",
                "The workflow was cancelled before publication; no staged artifact was retained.",
                OfficeWorkflowDiagnosticSeverity.Information,
                "cancel"));
            return new OfficeWorkflowResult(
                request.Id,
                request.Operation,
                OfficeWorkflowStatus.Cancelled,
                outputPath: null,
                inputBytes,
                outputBytes: 0,
                stopwatch.Elapsed,
                "Cancelled",
                diagnostics);
        } catch (Exception ex) when (ex is not OutOfMemoryException and not StackOverflowException) {
            diagnostics.Add(new OfficeWorkflowDiagnostic(
                "WorkflowFailed",
                ex.Message,
                OfficeWorkflowDiagnosticSeverity.Error,
                "execute",
                new Dictionary<string, string>(StringComparer.Ordinal) {
                    ["exceptionType"] = ex.GetType().Name
                }));
            return new OfficeWorkflowResult(
                request.Id,
                request.Operation,
                OfficeWorkflowStatus.Failed,
                outputPath: null,
                inputBytes,
                outputBytes: 0,
                stopwatch.Elapsed,
                "Workflow failed: " + ex.Message,
                diagnostics);
        } finally {
            if (stagingPath is not null) TryDelete(stagingPath);
        }
    }

    /// <summary>Runs a batch sequentially so every request shares one predictable local resource budget.</summary>
    public async Task<IReadOnlyList<OfficeWorkflowResult>> RunBatchAsync(
        IEnumerable<OfficeWorkflowRequest> requests,
        IProgress<OfficeWorkflowProgress>? progress = null,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(requests);
        OfficeWorkflowRequest[] batch = requests.ToArray();
        var results = new List<OfficeWorkflowResult>(batch.Length);
        for (int i = 0; i < batch.Length; i++) {
            if (cancellationToken.IsCancellationRequested) break;
            OfficeWorkflowRequest request = batch[i];
            int batchIndex = i;
            var batchProgress = progress is null
                ? null
                : new InlineProgress<OfficeWorkflowProgress>(item => progress.Report(new OfficeWorkflowProgress(
                    item.RequestId,
                    item.Stage,
                    $"{batchIndex + 1} of {batch.Length} · {item.Message}",
                    item.Fraction,
                    (batchIndex + item.Fraction) / Math.Max(1, batch.Length))));
            results.Add(await RunAsync(request, batchProgress, cancellationToken).ConfigureAwait(false));
        }
        return results;
    }

    private static OperationArtifact Execute(
        ValidatedRequest request,
        List<OfficeWorkflowDiagnostic> diagnostics,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        return request.Operation switch {
            OfficeWorkflowOperation.Convert => Convert(request, diagnostics, cancellationToken),
            OfficeWorkflowOperation.Inspect => Inspect(request, cancellationToken),
            OfficeWorkflowOperation.Compare => Compare(request, cancellationToken),
            OfficeWorkflowOperation.Optimize => Optimize(request, diagnostics, cancellationToken),
            OfficeWorkflowOperation.RepairPlan => RepairPlan(request, cancellationToken),
            OfficeWorkflowOperation.Repair => Repair(request, diagnostics, cancellationToken),
            OfficeWorkflowOperation.Sanitize => Sanitize(request, diagnostics, cancellationToken),
            _ => throw new ArgumentOutOfRangeException(nameof(request.Operation), request.Operation, "Unsupported workflow operation.")
        };
    }

    private static OperationArtifact Inspect(ValidatedRequest request, CancellationToken cancellationToken) {
        byte[] input = File.ReadAllBytes(request.InputPath);
        cancellationToken.ThrowIfCancellationRequested();
        PdfHealthSnapshot before = CreateHealthSnapshot(input, request.PdfReadOptions);
        cancellationToken.ThrowIfCancellationRequested();
        var report = new PdfHealthReport(
            OfficeWorkflowOperation.Inspect,
            before,
            after: null,
            before.CanRead ? "PDF inspection completed." : "PDF inspection found read blockers.",
            before.CanRead);
        return new OperationArtifact(null, report.Summary, report);
    }

    private static OperationArtifact Compare(ValidatedRequest request, CancellationToken cancellationToken) {
        byte[] expected = File.ReadAllBytes(request.InputPath);
        byte[] actual = File.ReadAllBytes(request.ComparisonPath!);
        PdfHealthSnapshot before = CreateHealthSnapshot(expected, request.PdfReadOptions);
        PdfHealthSnapshot after = CreateHealthSnapshot(actual, request.PdfReadOptions);
        PdfVisualComparisonReport comparison = PdfVisualComparer.Compare(
            expected,
            actual,
            cancellationToken,
            expectedReadOptions: request.PdfReadOptions,
            actualReadOptions: request.PdfReadOptions);
        var metrics = new Dictionary<string, string>(StringComparer.Ordinal) {
            ["pagesCompared"] = comparison.Pages.Count.ToString(System.Globalization.CultureInfo.InvariantCulture),
            ["differentPages"] = comparison.Pages.Count(page => !page.IsMatch).ToString(System.Globalization.CultureInfo.InvariantCulture),
            ["structuralDifferences"] = comparison.StructuralDifferences.Count.ToString(System.Globalization.CultureInfo.InvariantCulture)
        };
        string summary = comparison.IsMatch
            ? "The PDFs match within the managed structural and visual thresholds."
            : "The PDFs differ; review the structural findings and visual comparison gallery.";
        var report = new PdfHealthReport(
            OfficeWorkflowOperation.Compare,
            before,
            after,
            summary,
            comparison.IsMatch,
            metrics);
        byte[]? gallery = request.OutputPath is null
            ? null
            : Encoding.UTF8.GetBytes(comparison.ToHtmlGallery("OfficeIMO document comparison"));
        return new OperationArtifact(gallery, summary, report);
    }

    private static OperationArtifact Optimize(
        ValidatedRequest request,
        List<OfficeWorkflowDiagnostic> diagnostics,
        CancellationToken cancellationToken) {
        byte[] input = File.ReadAllBytes(request.InputPath);
        cancellationToken.ThrowIfCancellationRequested();
        PdfHealthSnapshot before = CreateHealthSnapshot(input, request.PdfReadOptions);
        cancellationToken.ThrowIfCancellationRequested();
        PdfOptimizationOptions options = PdfOptimizationOptions.Create(ToOptimizationProfile(request.OutputProfile));
        options.CancellationToken = cancellationToken;
        PdfOptimizationActionResult optimization = PdfDocument.Open(input, request.PdfReadOptions)
            .Optimization.Apply(options);
        cancellationToken.ThrowIfCancellationRequested();
        if (!optimization.PreservationReport.IsPreserved) {
            throw new InvalidOperationException("Lossless optimization preservation verification failed; no artifact will be published.");
        }
        byte[] output = optimization.Bytes;
        PdfHealthSnapshot after = CreateHealthSnapshot(output, request.PdfReadOptions);
        var metrics = new Dictionary<string, string>(StringComparer.Ordinal) {
            ["originalBytes"] = optimization.OriginalLengthBytes.ToString(System.Globalization.CultureInfo.InvariantCulture),
            ["candidateBytes"] = optimization.CandidateLengthBytes.ToString(System.Globalization.CultureInfo.InvariantCulture),
            ["returnedBytes"] = optimization.OptimizedLengthBytes.ToString(System.Globalization.CultureInfo.InvariantCulture),
            ["savedBytes"] = optimization.SavedBytes.ToString(System.Globalization.CultureInfo.InvariantCulture),
            ["actions"] = optimization.ActionCount.ToString(System.Globalization.CultureInfo.InvariantCulture),
            ["returnedOriginal"] = optimization.ReturnedOriginal.ToString()
        };
        diagnostics.Add(new OfficeWorkflowDiagnostic(
            "LosslessOptimization",
            optimization.ReturnedOriginal
                ? "The optimized candidate was not smaller, so the original bytes were retained."
                : $"Applied {optimization.ActionCount:N0} deterministic lossless optimization actions.",
            stage: "optimize",
            details: metrics));
        var report = new PdfHealthReport(
            OfficeWorkflowOperation.Optimize,
            before,
            after,
            optimization.ReturnedOriginal ? "The original was already the smaller lossless artifact." : $"Lossless optimization saved {optimization.SavedBytes:N0} bytes.",
            optimization.PreservationReport.IsPreserved,
            metrics);
        return new OperationArtifact(output, report.Summary, report);
    }

    private static OperationArtifact Repair(
        ValidatedRequest request,
        List<OfficeWorkflowDiagnostic> diagnostics,
        CancellationToken cancellationToken) {
        byte[] input = File.ReadAllBytes(request.InputPath);
        cancellationToken.ThrowIfCancellationRequested();
        PdfHealthSnapshot before = CreateHealthSnapshot(input, request.PdfReadOptions);
        cancellationToken.ThrowIfCancellationRequested();
        PdfRepairArtifactResult repair = PdfRepairArtifact.Create(
            input,
            new PdfRepairArtifactOptions {
                MaximumOutputBytes = request.Limits.MaximumOutputBytes,
                CancellationToken = cancellationToken
            },
            request.PdfReadOptions);
        cancellationToken.ThrowIfCancellationRequested();
        if (!repair.IsVerified) {
            throw new InvalidOperationException("Repair artifact verification failed; no artifact will be published.");
        }
        byte[] output = repair.ToBytes();
        PdfHealthSnapshot after = CreateHealthSnapshot(output, request.PdfReadOptions);
        var metrics = new Dictionary<string, string>(StringComparer.Ordinal) {
            ["recoveredDefects"] = repair.SourceRepairReport.RepairCount.ToString(System.Globalization.CultureInfo.InvariantCulture),
            ["detectedOnlyDefects"] = repair.SourceRepairReport.DetectionOnlyCount.ToString(System.Globalization.CultureInfo.InvariantCulture),
            ["strictOutputRepairs"] = repair.StrictOutputRepairReport.Diagnostics.Count.ToString(System.Globalization.CultureInfo.InvariantCulture),
            ["preserved"] = repair.Preservation.IsPreserved.ToString()
        };
        diagnostics.Add(new OfficeWorkflowDiagnostic(
            "VerifiedRepair",
            $"Persisted {repair.SourceRepairReport.RepairCount:N0} explicitly recovered structural defect(s), then reopened the artifact in strict mode.",
            stage: "repair",
            details: metrics));
        var report = new PdfHealthReport(
            OfficeWorkflowOperation.Repair,
            before,
            after,
            $"Created a verified normalized PDF from {repair.SourceRepairReport.RepairCount:N0} recovered defect(s).",
            repair.IsVerified,
            metrics);
        return new OperationArtifact(output, report.Summary, report);
    }

    private static OperationArtifact RepairPlan(ValidatedRequest request, CancellationToken cancellationToken) {
        byte[] input = File.ReadAllBytes(request.InputPath);
        cancellationToken.ThrowIfCancellationRequested();
        PdfHealthSnapshot before = CreateHealthSnapshot(input, request.PdfReadOptions);
        cancellationToken.ThrowIfCancellationRequested();
        PdfDocument document = PdfDocument.Open(input, request.PdfReadOptions);
        PdfAnalysisReport analysis = document.Analyze();
        PdfMutationPlan mutationPlan = document.PlanMutation(PdfMutationOperation.Optimize);
        PdfDocumentSecurityInfo security = analysis.Info.Security;
        bool hasProtectedSecurity = security.HasEncryption || security.HasSignatures ||
                                    security.HasDocMDPPermissions || security.HasUsageRights;
        bool canPersist = analysis.Repair.RepairCount > 0 &&
                          analysis.Repair.DetectionOnlyCount == 0 &&
                          mutationPlan.CanExecute &&
                          !hasProtectedSecurity;
        string blockers = string.Join(",", mutationPlan.BlockerCodes);
        var metrics = new Dictionary<string, string>(StringComparer.Ordinal) {
            ["recoveredDefects"] = analysis.Repair.RepairCount.ToString(System.Globalization.CultureInfo.InvariantCulture),
            ["detectedOnlyDefects"] = analysis.Repair.DetectionOnlyCount.ToString(System.Globalization.CultureInfo.InvariantCulture),
            ["canonicalMutationMode"] = mutationPlan.ExecutionMode.ToString(),
            ["canonicalBlockers"] = blockers,
            ["protectedSecurity"] = hasProtectedSecurity.ToString(),
            ["canCreateRepairArtifact"] = canPersist.ToString()
        };
        string summary = canPersist
            ? $"Repair plan is executable: {analysis.Repair.RepairCount:N0} recovered defect(s) can be persisted and verified."
            : analysis.Repair.RepairCount == 0
                ? "Repair plan is not needed: no recovered structural defects were reported."
                : "Repair plan is blocked by the canonical rewrite, defect, or security policy; review the plan evidence before creating an artifact.";
        var report = new PdfHealthReport(
            OfficeWorkflowOperation.RepairPlan,
            before,
            after: null,
            summary,
            canPersist,
            metrics);
        return new OperationArtifact(null, summary, report);
    }

    private static OperationArtifact Sanitize(
        ValidatedRequest request,
        List<OfficeWorkflowDiagnostic> diagnostics,
        CancellationToken cancellationToken) {
        byte[] input = File.ReadAllBytes(request.InputPath);
        cancellationToken.ThrowIfCancellationRequested();
        PdfHealthSnapshot before = CreateHealthSnapshot(input, request.PdfReadOptions);
        cancellationToken.ThrowIfCancellationRequested();
        PdfSanitizationResult sanitization = PdfDocument.Open(input, request.PdfReadOptions).Sanitize(
            new PdfSanitizationOptions { CancellationToken = cancellationToken });
        cancellationToken.ThrowIfCancellationRequested();
        if (!sanitization.IsSanitized || !sanitization.PreservationReport.IsPreserved) {
            throw new InvalidOperationException("Sanitization verification failed; no artifact will be published.");
        }
        byte[] output = sanitization.ToBytes();
        PdfHealthSnapshot after = CreateHealthSnapshot(output, request.PdfReadOptions);
        var metrics = new Dictionary<string, string>(StringComparer.Ordinal) {
            ["removedFindings"] = sanitization.RemovedFindings.Count.ToString(System.Globalization.CultureInfo.InvariantCulture),
            ["remainingFindings"] = sanitization.RemainingFindings.Count.ToString(System.Globalization.CultureInfo.InvariantCulture),
            ["quarantinedAttachments"] = sanitization.QuarantinedAttachments.Count.ToString(System.Globalization.CultureInfo.InvariantCulture),
            ["preserved"] = sanitization.PreservationReport.IsPreserved.ToString()
        };
        diagnostics.Add(new OfficeWorkflowDiagnostic(
            "SanitizationProof",
            $"Removed {sanitization.RemovedFindings.Count:N0} forbidden item(s); post-save inventory found {sanitization.RemainingFindings.Count:N0} remaining.",
            stage: "sanitize",
            details: metrics));
        var report = new PdfHealthReport(
            OfficeWorkflowOperation.Sanitize,
            before,
            after,
            $"Sanitization removed {sanitization.RemovedFindings.Count:N0} forbidden item(s) and verified the saved artifact.",
            sanitization.IsSanitized && sanitization.PreservationReport.IsPreserved,
            metrics);
        return new OperationArtifact(output, report.Summary, report);
    }

    private static PdfHealthSnapshot CreateHealthSnapshot(byte[] bytes, PdfReadOptions readOptions) {
        PdfDocumentPreflight preflight = PdfDocument.Preflight(bytes, readOptions);
        PdfDocumentInfo? info = preflight.DocumentInfo;
        PdfRepairReport? repairs = null;
        var diagnostics = new List<string>(preflight.Diagnostics);
        if (preflight.CanRead) {
            try {
                PdfAnalysisReport analysis = PdfDocument.Open(bytes, readOptions).Analyze();
                info = analysis.Info;
                repairs = analysis.Repair;
                diagnostics.AddRange(repairs.Diagnostics.Select(static item => item.Message));
            } catch (Exception ex) when (ex is not OutOfMemoryException and not StackOverflowException) {
                diagnostics.Add("Analysis: " + ex.Message);
            }
        }

        PdfDocumentSecurityInfo security = info?.Security ?? preflight.Probe.Security;
        return new PdfHealthSnapshot(
            bytes.LongLength,
            info?.PageCount ?? 0,
            info?.EffectiveVersion ?? info?.HeaderVersion,
            preflight.CanRead,
            preflight.CanRewrite,
            security.HasEncryption,
            security.HasSignatures,
            info?.HasTaggedContent == true,
            info?.HasActiveContent == true,
            info?.HasEmbeddedFiles == true,
            repairs?.RepairCount ?? 0,
            repairs?.DetectionOnlyCount ?? 0,
            diagnostics);
    }

    private static ValidatedRequest ValidateRequest(OfficeWorkflowRequest request) {
        if (string.IsNullOrWhiteSpace(request.Id)) throw new ArgumentException("Request id cannot be empty.", nameof(request));
        if (string.IsNullOrWhiteSpace(request.InputPath)) throw new ArgumentException("Input path cannot be empty.", nameof(request));
        string inputPath = Path.GetFullPath(request.InputPath);
        if (!File.Exists(inputPath)) throw new FileNotFoundException("The workflow input file does not exist.", inputPath);
        OfficeWorkflowLimits limits = (request.Limits ?? throw new ArgumentException("Workflow limits cannot be null.", nameof(request))).CloneAndValidate();
        OfficeWorkflowRoute? route = null;
        string? comparisonPath = null;
        string? outputPath = string.IsNullOrWhiteSpace(request.OutputPath) ? null : Path.GetFullPath(request.OutputPath);

        if (request.Operation == OfficeWorkflowOperation.Convert) {
            route = OfficeWorkflowCatalog.Find(request.ConversionRouteId)
                ?? throw new ArgumentException("Choose a supported conversion route.", nameof(request));
            string extension = Path.GetExtension(inputPath);
            if (!route.SourceExtensions.Any(item => string.Equals(NormalizeExtension(item), extension, StringComparison.OrdinalIgnoreCase))) {
                throw new ArgumentException($"Route '{route.Id}' does not accept '{extension}' input.", nameof(request));
            }
            outputPath ??= Path.ChangeExtension(inputPath, NormalizeExtension(route.TargetExtension));
            if (!string.Equals(Path.GetExtension(outputPath), NormalizeExtension(route.TargetExtension), StringComparison.OrdinalIgnoreCase)) {
                throw new ArgumentException($"Route '{route.Id}' requires a '{NormalizeExtension(route.TargetExtension)}' output.", nameof(request));
            }
        } else if (request.Operation == OfficeWorkflowOperation.Compare) {
            if (string.IsNullOrWhiteSpace(request.ComparisonPath)) throw new ArgumentException("PDF comparison requires a second input path.", nameof(request));
            comparisonPath = Path.GetFullPath(request.ComparisonPath);
            if (!File.Exists(comparisonPath)) throw new FileNotFoundException("The comparison PDF does not exist.", comparisonPath);
            EnsurePdfExtension(inputPath);
            EnsurePdfExtension(comparisonPath);
            if (outputPath is not null && !string.Equals(Path.GetExtension(outputPath), ".html", StringComparison.OrdinalIgnoreCase)) {
                throw new ArgumentException("Comparison output must be an HTML gallery.", nameof(request));
            }
        } else {
            EnsurePdfExtension(inputPath);
            if (request.Operation is OfficeWorkflowOperation.Optimize or OfficeWorkflowOperation.Repair or OfficeWorkflowOperation.Sanitize) {
                outputPath ??= Path.Combine(
                    Path.GetDirectoryName(inputPath)!,
                    Path.GetFileNameWithoutExtension(inputPath) + "." + request.Operation.ToString().ToLowerInvariant() + ".pdf");
                EnsurePdfExtension(outputPath);
            } else if (request.Operation is OfficeWorkflowOperation.Inspect or OfficeWorkflowOperation.RepairPlan && outputPath is not null) {
                throw new ArgumentException("The selected report-only operation does not publish an artifact.", nameof(request));
            }
        }

        return new ValidatedRequest(
            request.Id,
            request.Operation,
            inputPath,
            comparisonPath,
            outputPath,
            route,
            request.ConflictPolicy,
            request.OutputProfile,
            limits,
            new PdfReadOptions { Password = request.PdfPassword });
    }

    private static void ValidateStagedArtifact(string stagingPath, string outputPath, PdfReadOptions readOptions) {
        string extension = Path.GetExtension(outputPath).ToLowerInvariant();
        switch (extension) {
            case ".pdf":
                PdfDocumentInfo info = PdfDocument.Open(stagingPath, readOptions).Inspect();
                if (info.PageCount == 0) throw new InvalidOperationException("Generated PDF has no pages.");
                break;
            case ".docx":
                using (FileStream stream = File.OpenRead(stagingPath))
                using (WordDocument document = WordDocument.Load(stream)) { }
                break;
            case ".xlsx":
                using (FileStream stream = File.OpenRead(stagingPath))
                using (ExcelDocument document = ExcelDocument.Load(stream)) { }
                break;
            case ".pptx":
                using (FileStream stream = File.OpenRead(stagingPath))
                using (PowerPointPresentation document = PowerPointPresentation.Load(stream)) { }
                break;
            case ".html": {
                string html = File.ReadAllText(stagingPath, Encoding.UTF8);
                _ = HtmlConversionDocument.Parse(html);
                break;
            }
            default:
                throw new NotSupportedException("No output validator is registered for '" + extension + "'.");
        }
    }

    private static string Publish(string stagingPath, string requestedPath, OfficeWorkflowConflictPolicy policy) {
        switch (policy) {
            case OfficeWorkflowConflictPolicy.Fail:
                File.Move(stagingPath, requestedPath, overwrite: false);
                return requestedPath;
            case OfficeWorkflowConflictPolicy.Replace:
                File.Move(stagingPath, requestedPath, overwrite: true);
                return requestedPath;
            case OfficeWorkflowConflictPolicy.Rename:
                for (int suffix = 0; suffix < 10_000; suffix++) {
                    string candidate = suffix == 0 ? requestedPath : AddSuffix(requestedPath, suffix);
                    try {
                        File.Move(stagingPath, candidate, overwrite: false);
                        return candidate;
                    } catch (IOException) when (File.Exists(candidate)) {
                        // Another request owns this candidate. Try the next deterministic suffix.
                    }
                }
                throw new IOException("No available numbered output path could be reserved.");
            default:
                throw new ArgumentOutOfRangeException(nameof(policy), policy, "Unsupported conflict policy.");
        }
    }

    private static void EnsureVerifiedHealthArtifact(OfficeWorkflowOperation operation, PdfHealthReport? report) {
        if (operation is OfficeWorkflowOperation.Optimize or OfficeWorkflowOperation.Repair or OfficeWorkflowOperation.Sanitize &&
            report is not { Verified: true }) {
            throw new InvalidOperationException($"{operation} did not produce verified preservation evidence; no artifact will be published.");
        }
    }

    private static string AddSuffix(string path, int suffix) => Path.Combine(
        Path.GetDirectoryName(path)!,
        Path.GetFileNameWithoutExtension(path) + " (" + suffix.ToString(System.Globalization.CultureInfo.InvariantCulture) + ")" + Path.GetExtension(path));

    private static void AddPdfWarnings(IEnumerable<PdfConversionWarning> warnings, List<OfficeWorkflowDiagnostic> diagnostics) {
        foreach (PdfConversionWarning warning in warnings) {
            diagnostics.Add(new OfficeWorkflowDiagnostic(
                warning.Code,
                warning.Message,
                warning.Severity == PdfConversionWarningSeverity.Information
                    ? OfficeWorkflowDiagnosticSeverity.Information
                    : OfficeWorkflowDiagnosticSeverity.Warning,
                "convert"));
        }
    }

    private static void AddMessages(IEnumerable<string> warnings, bool hasLoss, List<OfficeWorkflowDiagnostic> diagnostics) {
        foreach (string warning in warnings) {
            diagnostics.Add(new OfficeWorkflowDiagnostic(
                "ConversionWarning",
                warning,
                hasLoss ? OfficeWorkflowDiagnosticSeverity.Warning : OfficeWorkflowDiagnosticSeverity.Information,
                "convert"));
        }
    }

    private static PdfExportProfile ToPdfExportProfile(OfficeWorkflowOutputProfile profile) => profile switch {
        OfficeWorkflowOutputProfile.Faithful => PdfExportProfile.Faithful,
        OfficeWorkflowOutputProfile.Lightweight => PdfExportProfile.Lightweight,
        OfficeWorkflowOutputProfile.PrintReady => PdfExportProfile.PrintReady,
        OfficeWorkflowOutputProfile.TextOnly => PdfExportProfile.TextOnly,
        _ => throw new ArgumentOutOfRangeException(nameof(profile), profile, "Unsupported output profile.")
    };

    private static PdfOptimizationProfile ToOptimizationProfile(OfficeWorkflowOutputProfile profile) => profile switch {
        OfficeWorkflowOutputProfile.Faithful => PdfOptimizationProfile.Balanced,
        OfficeWorkflowOutputProfile.Lightweight => PdfOptimizationProfile.MaximumCompression,
        OfficeWorkflowOutputProfile.PrintReady => PdfOptimizationProfile.Archival,
        OfficeWorkflowOutputProfile.TextOnly => PdfOptimizationProfile.Web,
        _ => throw new ArgumentOutOfRangeException(nameof(profile), profile, "Unsupported output profile.")
    };

    private static byte[] SerializePdfConversion(PdfDocumentConversionResult conversion, CancellationToken cancellationToken) {
        using var stream = new MemoryStream();
        conversion.SaveAsync(stream, cancellationToken).GetAwaiter().GetResult().RequireSuccess();
        cancellationToken.ThrowIfCancellationRequested();
        return stream.ToArray();
    }

    private static PdfPowerPointImportOptions CreatePowerPointImportOptions(CancellationToken cancellationToken) {
        PdfPowerPointImportOptions options = PdfPowerPointImportOptions.CreateEditableContent();
        options.CancellationToken = cancellationToken;
        return options;
    }

    private static OfficeWorkflowResult CreateResult(
        ValidatedRequest request,
        OfficeWorkflowStatus status,
        string? outputPath,
        long inputBytes,
        long outputBytes,
        TimeSpan duration,
        string summary,
        IReadOnlyList<OfficeWorkflowDiagnostic> diagnostics,
        PdfHealthReport? report) => new(
            request.Id,
            request.Operation,
            status,
            outputPath,
            inputBytes,
            outputBytes,
            duration,
            summary,
            diagnostics,
            report);

    private static void Report(IProgress<OfficeWorkflowProgress>? progress, string id, string stage, string message, double fraction) =>
        progress?.Report(new OfficeWorkflowProgress(id, stage, message, fraction));

    private static void EnforceInputLimit(string path, long size, OfficeWorkflowLimits limits) {
        if (size > limits.MaximumInputBytes) {
            throw new InvalidOperationException(
                $"Input '{Path.GetFileName(path)}' is {size:N0} bytes, above the configured {limits.MaximumInputBytes:N0}-byte limit.");
        }
    }

    private static void EnsurePdfExtension(string path) {
        if (!string.Equals(Path.GetExtension(path), ".pdf", StringComparison.OrdinalIgnoreCase)) {
            throw new ArgumentException("Document Health accepts PDF inputs.", nameof(path));
        }
    }

    private static string NormalizeExtension(string extension) => extension.StartsWith('.') ? extension : "." + extension;

    private static string DescribeOperation(OfficeWorkflowOperation operation) => operation switch {
        OfficeWorkflowOperation.Convert => "Converting with the first-party OfficeIMO format owner",
        OfficeWorkflowOperation.Inspect => "Inspecting PDF structure and capabilities",
        OfficeWorkflowOperation.Compare => "Comparing PDF structure and managed render output",
        OfficeWorkflowOperation.Optimize => "Applying deterministic lossless PDF optimization",
        OfficeWorkflowOperation.RepairPlan => "Planning a verified PDF repair artifact",
        OfficeWorkflowOperation.Repair => "Creating a verified PDF repair artifact",
        OfficeWorkflowOperation.Sanitize => "Removing forbidden PDF content and verifying the result",
        _ => "Running workflow"
    };

    private static void TryDelete(string path) {
        try {
            if (File.Exists(path)) File.Delete(path);
        } catch (IOException) {
            // Best-effort cleanup; the random staging name is never returned or published.
        } catch (UnauthorizedAccessException) {
            // Best-effort cleanup; the random staging name is never returned or published.
        }
    }

    private sealed record OperationArtifact(byte[]? Bytes, string Summary, PdfHealthReport? HealthReport);

    private sealed class InlineProgress<T>(Action<T> report) : IProgress<T> {
        public void Report(T value) => report(value);
    }

    private sealed record ValidatedRequest(
        string Id,
        OfficeWorkflowOperation Operation,
        string InputPath,
        string? ComparisonPath,
        string? OutputPath,
        OfficeWorkflowRoute? Route,
        OfficeWorkflowConflictPolicy ConflictPolicy,
        OfficeWorkflowOutputProfile OutputProfile,
        OfficeWorkflowLimits Limits,
        PdfReadOptions PdfReadOptions);
}
