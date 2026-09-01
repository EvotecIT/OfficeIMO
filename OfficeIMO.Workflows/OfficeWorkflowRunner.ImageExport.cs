using System.Diagnostics;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRunner : IOfficeOutputWorkflowRunner {
    /// <inheritdoc />
    public async Task<PdfPageImageExportResult> ExportPdfPagesAsync(
        PdfPageImageExportRequest request,
        IProgress<OfficeWorkflowProgress>? progress = null,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(request);
        var stopwatch = Stopwatch.StartNew();
        var diagnostics = new List<OfficeWorkflowDiagnostic>();
        string? stagingDirectory = null;
        long inputBytes = 0L;
        WorkflowFailureStage failureStage = WorkflowFailureStage.Validation;

        try {
            ValidatedImageExportRequest validated = ValidateImageExportRequest(request);
            failureStage = WorkflowFailureStage.Input;
            inputBytes = new FileInfo(validated.InputPath).Length;
            EnforceInputLimit(validated.InputPath, inputBytes, validated.Limits);
            Report(progress, validated.Id, "validate", "Validating PDF and page selection", 0.05D);
            cancellationToken.ThrowIfCancellationRequested();

            PdfDocument document = PdfDocument.Load(validated.InputPath, validated.LoadOptions);
            PdfDocumentInfo info = document.Inspect();
            int[] pageNumbers = ResolvePageNumbers(validated.Pages, info.PageCount);
            if (pageNumbers.Length > validated.MaximumPages) {
                throw new InvalidOperationException(
                    $"The selection contains {pageNumbers.Length:N0} pages, above the configured {validated.MaximumPages:N0}-page limit.");
            }

            string parent = Path.GetDirectoryName(validated.OutputDirectory)!;
            Directory.CreateDirectory(parent);
            stagingDirectory = Path.Combine(parent, "." + Path.GetFileName(validated.OutputDirectory) + "." + Guid.NewGuid().ToString("N") + ".tmp");

            var options = new PdfImageExportOptions {
                TargetDpi = validated.TargetDpi,
                ThumbnailMaxDimension = validated.MaximumDimension,
                MaximumOutputCount = validated.MaximumPages,
                MaximumTotalEncodedBytes = validated.Limits.MaximumOutputBytes
            };
            var exportProgress = new Progress<OfficeImageExportProgress>(item => {
                double fraction = pageNumbers.Length == 0 ? 0D : (double)item.CompletedCount / pageNumbers.Length;
                Report(progress, validated.Id, "render", "Rendering selected PDF pages", 0.12D + fraction * 0.58D);
            });
            options.Progress = exportProgress;

            PdfDocumentImageExportBuilder builder = document
                .ToImages(options)
                .Pages(PdfPageSelection.From(pageNumbers))
                .As(validated.Format);
            Report(progress, validated.Id, "render", "Rendering selected PDF pages", 0.12D);
            OfficeImageExportBatchSaveResult saved = await builder
                .SaveFilesAsync(stagingDirectory, cancellationToken)
                .ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
            failureStage = WorkflowFailureStage.Output;

            long outputBytes = 0L;
            for (int index = 0; index < saved.Files.Count; index++) {
                OfficeImageExportSavedFile file = saved.Files[index];
                outputBytes = checked(outputBytes + file.EncodedLength);
                byte[] bytes = await File.ReadAllBytesAsync(file.Path, cancellationToken).ConfigureAwait(false);
                if (!OfficeImageReader.TryValidateContent(bytes, file.Path, out OfficeImageInfo imageInfo) ||
                    imageInfo.Width != file.Width || imageInfo.Height != file.Height) {
                    throw new InvalidOperationException("A staged page image failed content and dimension validation.");
                }
                AddImageDiagnostics(file.Diagnostics, diagnostics, pageNumbers[index]);
            }
            if (outputBytes > validated.Limits.MaximumOutputBytes) {
                throw new InvalidOperationException(
                    $"Generated images total {outputBytes:N0} bytes, above the configured {validated.Limits.MaximumOutputBytes:N0}-byte limit.");
            }

            diagnostics.Add(new OfficeWorkflowDiagnostic(
                "PageImagesReopened",
                "Every staged page image was decoded and its dimensions were verified before publication.",
                stage: "validate-output",
                details: new Dictionary<string, string>(StringComparer.Ordinal) {
                    ["pageCount"] = saved.Files.Count.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    ["format"] = validated.Format.ToString(),
                    ["outputBytes"] = outputBytes.ToString(System.Globalization.CultureInfo.InvariantCulture)
                }));
            Report(progress, validated.Id, "publish", "Publishing the validated image folder", 0.9D);
            string publishedDirectory = await PublishDirectoryAsync(
                    stagingDirectory,
                    validated.OutputDirectory,
                    validated.ConflictPolicy,
                    diagnostics,
                    cancellationToken)
                .ConfigureAwait(false);
            stagingDirectory = null;

            var files = new List<PdfPageImageFile>(saved.Files.Count);
            for (int index = 0; index < saved.Files.Count; index++) {
                OfficeImageExportSavedFile file = saved.Files[index];
                files.Add(new PdfPageImageFile(
                    pageNumbers[index],
                    Path.Combine(publishedDirectory, Path.GetFileName(file.Path)),
                    file.Format,
                    file.Width,
                    file.Height,
                    file.EncodedLength));
            }
            Report(progress, validated.Id, "complete", "Page images are ready", 1D);
            return new PdfPageImageExportResult(
                validated.Id,
                OfficeWorkflowStatus.Completed,
                OfficeWorkflowFailureKind.None,
                publishedDirectory,
                inputBytes,
                outputBytes,
                stopwatch.Elapsed,
                $"Exported {files.Count:N0} PDF {(files.Count == 1 ? "page" : "pages")} as {validated.Format}.",
                files,
                diagnostics);
        } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
            diagnostics.Add(new OfficeWorkflowDiagnostic(
                "Cancelled",
                "Page export was cancelled before publication.",
                OfficeWorkflowDiagnosticSeverity.Information,
                "cancel"));
            return new PdfPageImageExportResult(
                request.Id,
                OfficeWorkflowStatus.Cancelled,
                OfficeWorkflowFailureKind.None,
                null,
                inputBytes,
                0L,
                stopwatch.Elapsed,
                "Cancelled",
                Array.Empty<PdfPageImageFile>(),
                diagnostics);
        } catch (Exception ex) when (ex is not OutOfMemoryException and not StackOverflowException) {
            IReadOnlyDictionary<string, string> details = CreateFailureDetails(ex);
            diagnostics.Add(new OfficeWorkflowDiagnostic(
                "PageImageExportFailed",
                ex.Message,
                OfficeWorkflowDiagnosticSeverity.Error,
                "execute",
                details));
            return new PdfPageImageExportResult(
                request.Id,
                OfficeWorkflowStatus.Failed,
                ClassifyFailure(ex, failureStage),
                null,
                inputBytes,
                0L,
                stopwatch.Elapsed,
                "Page image export failed: " + ex.Message,
                Array.Empty<PdfPageImageFile>(),
                diagnostics);
        } finally {
            if (stagingDirectory is not null) TryDeleteDirectory(stagingDirectory);
        }
    }

    private static ValidatedImageExportRequest ValidateImageExportRequest(PdfPageImageExportRequest request) {
        if (string.IsNullOrWhiteSpace(request.Id)) throw new ArgumentException("Request id cannot be empty.", nameof(request));
        if (string.IsNullOrWhiteSpace(request.InputPath)) throw new ArgumentException("Input path cannot be empty.", nameof(request));
        if (string.IsNullOrWhiteSpace(request.OutputDirectory)) throw new ArgumentException("Output directory cannot be empty.", nameof(request));
        string inputPath = Path.GetFullPath(request.InputPath);
        if (!File.Exists(inputPath)) throw new FileNotFoundException("The source PDF does not exist.", inputPath);
        EnsurePdfExtension(inputPath);
        string outputDirectory = Path.GetFullPath(request.OutputDirectory);
        if (OfficeWorkflowPathIdentity.AreEquivalent(inputPath, outputDirectory)) {
            throw new ArgumentException("Output directory cannot be the source PDF path.", nameof(request));
        }
        if (!Enum.IsDefined(request.Format)) throw new ArgumentOutOfRangeException(nameof(request.Format));
        if (request.TargetDpi <= 0D || double.IsNaN(request.TargetDpi) || double.IsInfinity(request.TargetDpi)) {
            throw new ArgumentOutOfRangeException(nameof(request.TargetDpi));
        }
        if (request.MaximumDimension is < 1) throw new ArgumentOutOfRangeException(nameof(request.MaximumDimension));
        if (request.MaximumPages < 1) throw new ArgumentOutOfRangeException(nameof(request.MaximumPages));
        if (!Enum.IsDefined(request.ConflictPolicy)) throw new ArgumentOutOfRangeException(nameof(request.ConflictPolicy));
        OfficeWorkflowLimits limits = (request.Limits ?? throw new ArgumentException("Workflow limits cannot be null.", nameof(request))).CloneAndValidate();
        return new ValidatedImageExportRequest(
            request.Id,
            inputPath,
            outputDirectory,
            request.Pages,
            request.Format,
            request.TargetDpi,
            request.MaximumDimension,
            request.MaximumPages,
            request.ConflictPolicy,
            limits,
            new PdfLoadOptions { Password = request.PdfPassword });
    }

    private static int[] ResolvePageNumbers(string? selector, int pageCount) {
        if (pageCount < 1) throw new InvalidOperationException("The source PDF has no pages.");
        if (string.IsNullOrWhiteSpace(selector)) return Enumerable.Range(1, pageCount).ToArray();
        return PdfPageSelector.Parse(selector)
            .ResolveSelection(pageCount)
            .Ranges
            .SelectMany(static range => Enumerable.Range(range.FirstPage, range.PageCount))
            .ToArray();
    }

    private static void AddImageDiagnostics(
        IReadOnlyList<OfficeImageExportDiagnostic> source,
        ICollection<OfficeWorkflowDiagnostic> destination,
        int pageNumber) {
        foreach (OfficeImageExportDiagnostic item in source) {
            OfficeWorkflowDiagnosticSeverity severity = item.Severity switch {
                OfficeImageExportDiagnosticSeverity.Error => OfficeWorkflowDiagnosticSeverity.Error,
                OfficeImageExportDiagnosticSeverity.Warning => OfficeWorkflowDiagnosticSeverity.Warning,
                _ => OfficeWorkflowDiagnosticSeverity.Information
            };
            destination.Add(new OfficeWorkflowDiagnostic(
                item.Code,
                item.Message,
                severity,
                "render",
                new Dictionary<string, string>(StringComparer.Ordinal) {
                    ["pageNumber"] = pageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture)
                }));
        }
    }

    private sealed record ValidatedImageExportRequest(
        string Id,
        string InputPath,
        string OutputDirectory,
        string? Pages,
        OfficeImageExportFormat Format,
        double TargetDpi,
        int? MaximumDimension,
        int MaximumPages,
        OfficeWorkflowConflictPolicy ConflictPolicy,
        OfficeWorkflowLimits Limits,
        PdfLoadOptions LoadOptions);
}
