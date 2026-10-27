namespace OfficeIMO;

public static partial class OfficeConverter {
    /// <summary>Converts a bounded selection using the same adapters as single-file conversion.</summary>
    public static Task<OfficeBatchReport> ConvertAsync(IEnumerable<string> sources, string outputDirectory,
        OfficeDocumentFormat target, OfficeBatchOptions? options = null, IProgress<OfficeBatchItemReport>? progress = null,
        CancellationToken cancellationToken = default) {
        if (sources == null) throw new ArgumentNullException(nameof(sources));
        OfficeBatchOptions settings = SelectBatchTarget(options ?? new(), outputDirectory, target);
        string[] paths = sources.Take(checked(settings.MaximumFiles + 1)).ToArray();
        return ConvertAsync(settings with { InputPaths = paths }, progress, cancellationToken);
    }

    /// <summary>Discovers and converts a directory incrementally.</summary>
    public static Task<OfficeBatchReport> ConvertDirectoryAsync(string sourceDirectory, string outputDirectory,
        OfficeDocumentFormat target, OfficeBatchOptions? options = null, IProgress<OfficeBatchItemReport>? progress = null,
        CancellationToken cancellationToken = default) => ConvertAsync(SelectBatchTarget(options ?? new(), outputDirectory, target)
            with { InputDirectory = sourceDirectory, InputPaths = null }, progress, cancellationToken);

    /// <summary>Runs a portable batch definition, retaining bounded observations and complete counts.</summary>
    public static async Task<OfficeBatchReport> ConvertAsync(OfficeBatchOptions options,
        IProgress<OfficeBatchItemReport>? progress = null, CancellationToken cancellationToken = default) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        var collector = new OfficeBatchReportCollector();
        DateTimeOffset started = DateTimeOffset.UtcNow;
        OfficeBatchOptions settings = options with { ConversionOptions = options.ConversionOptions.Snapshot() };
        if (settings.SourcePasswordEnvironmentVariable is { } reference)
            settings.ConversionOptions.SourcePassword = ResolveBatchSecret(settings, reference);
        if (settings.PdfPasswordEnvironmentVariable is { } pdfReference)
            settings.ConversionOptions.PdfLoad = (settings.ConversionOptions.PdfLoad ?? new Pdf.PdfLoadOptions()).Snapshot(null, ResolveBatchSecret(settings, pdfReference));
        var report = new OfficeBatchReport { Summary = new(0, 0, 0, 0, false), State = OfficeBatchRunState.Running,
            StartedUtc = started, RunId = settings.ReportPath == null ? null : Guid.NewGuid().ToString("N"),
            OutputDirectory = Path.GetFullPath(settings.OutputDirectory), ResumeFrom = settings.ResumeFrom,
            RetryFailed = settings.RetryFailed, AllowEmpty = settings.AllowEmpty,
            JobPath = settings.JobPath, JobSha256 = settings.JobSha256 };
        if (settings.JobPath is { } savedPath) {
            OfficeBatchOptions saved = OfficeBatchOptionsSerializer.ReadSnapshot(savedPath, out string hash) with { RetryFailed = settings.RetryFailed, ReportPath = options.ReportPath, AllowEmpty = settings.AllowEmpty };
            string directory = Path.GetDirectoryName(savedPath)!;
            if (hash != settings.JobSha256 || saved.ToJson(directory) != options.ToJson(directory))
                throw new InvalidDataException("The saved batch changed or execution settings differ from its definition. Save and review it before running.");
        }
        string? reportPath = settings.ReportPath == null ? null : Path.GetFullPath(settings.ReportPath);
        if (reportPath != null) ValidateBatchReportPath(settings, reportPath);
        using var recording = reportPath == null ? null : new RunRecording(reportPath, report.RunId!, settings.RetainReport, settings.PublicationGuard);
        if (recording != null) await recording.BeginAsync(report).ConfigureAwait(false);
        try {
            OfficeBatchSummary summary = await OfficeBatchExecutor.RunAsync(new OfficeConversionExecutor(), settings,
                new BatchProgress(collector, progress), cancellationToken, settings.PublicationGuard).ConfigureAwait(false);
            OfficeBatchReport observed = collector.Complete(summary, settings.AllowEmpty);
            report = report with { Summary = observed.Summary, Items = observed.Items, ItemsTruncated = observed.ItemsTruncated,
                Succeeded = observed.Succeeded, FinishedUtc = DateTimeOffset.UtcNow,
                State = summary.Cancelled ? OfficeBatchRunState.Cancelled : observed.Succeeded ? OfficeBatchRunState.Completed : OfficeBatchRunState.Failed };
        } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
            OfficeBatchReport observed = collector.Complete(collector.ObservedSummary(true), settings.AllowEmpty);
            report = report with { Summary = observed.Summary, Items = observed.Items, ItemsTruncated = observed.ItemsTruncated,
                FinishedUtc = DateTimeOffset.UtcNow, State = OfficeBatchRunState.Cancelled };
        } catch (Exception error) when (error is not OutOfMemoryException and not StackOverflowException) {
            string failure = "Batch execution failed (" + error.GetType().Name + "). Review the original execution error before retrying.";
            OfficeBatchReport observed = collector.Complete(collector.ObservedSummary(false, failure), settings.AllowEmpty);
            report = report with { Summary = observed.Summary, Items = observed.Items, ItemsTruncated = observed.ItemsTruncated,
                FinishedUtc = DateTimeOffset.UtcNow, State = OfficeBatchRunState.Failed, FailureSummary = failure };
            if (recording != null) _ = await recording.FinishAsync(report).ConfigureAwait(false);
            throw;
        }
        return recording == null ? report : await recording.FinishAsync(report).ConfigureAwait(false);
    }

    /// <summary>Plans a batch without rendering, creating output or resolving runtime secrets.</summary>
    public static Task<OfficeBatchPlan> PlanAsync(OfficeBatchOptions options, IProgress<OfficeBatchPlanItem>? progress = null,
        CancellationToken cancellationToken = default) => OfficeBatchExecutor.PlanAsync(new OfficeConversionExecutor(), options, progress, cancellationToken);

    private static string ResolveBatchSecret(OfficeBatchOptions settings, string reference) {
        OfficeBatchJson.CheckEnvironmentName(reference);
        string? value = settings.ResolveEnvironmentVariable?.Invoke(reference);
        if (value == null || value.Length == 0) throw new InvalidOperationException("The batch requires an admitted runtime secret: " + reference);
        return value;
    }

    private static void ValidateBatchReportPath(OfficeBatchOptions settings, string path) {
        OfficeBatchExecutor.EnsureNoLinks(path); OfficeBatchExecutor.EnsureNoLinks(path + ".lease");
        foreach (string root in new[] { settings.InputDirectory, settings.OutputDirectory, settings.CheckpointDirectory, settings.JobPath }.OfType<string>())
            if (OfficeIMO.Internal.OfficePathIdentity.IsSameOrDescendant(path, root) || OfficeIMO.Internal.OfficePathIdentity.IsSameOrDescendant(root, path))
                throw new InvalidDataException("Save run reports outside the source, output, checkpoint and job paths.");
        foreach (string input in settings.InputPaths ?? [])
            if (OfficeIMO.Internal.OfficePathIdentity.IsSameOrDescendant(path, input)) throw new InvalidDataException("The run report must not replace a source file or package member.");
    }

    private static OfficeBatchOptions SelectBatchTarget(OfficeBatchOptions options, string outputDirectory, OfficeDocumentFormat target) {
        string extension = options.TargetExtension;
        if (extension == ".pdf" && target != OfficeDocumentFormat.Pdf) extension = DefaultExtension(target);
        if (FormatForExtension(extension) != target) throw new ArgumentException("The batch target extension must agree with the selected document format.");
        return options with { OutputDirectory = outputDirectory, TargetExtension = extension };
    }

    private sealed class BatchProgress(OfficeBatchReportCollector collector, IProgress<OfficeBatchItemReport>? next) : IProgress<OfficeBatchItemReport> {
        public void Report(OfficeBatchItemReport item) { collector.Report(item); next?.Report(item); }
    }
}
