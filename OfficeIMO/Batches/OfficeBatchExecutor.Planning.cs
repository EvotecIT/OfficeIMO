using OfficeIMO.Internal;

namespace OfficeIMO;

internal static partial class OfficeBatchExecutor {
    /// <summary>Collects the distinct routes in a bounded selection for host option binding, without rendering or writes.</summary>
    internal static IReadOnlyList<OfficeConversionCapability> SelectCapabilities(OfficeBatchOptions request, CancellationToken token) {
        var routes = OfficeConverter.ExecutionRoutes.Where(route => route.CanExecute).ToArray();
        var settings = Snapshot(request, routes);
        var selected = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        long count = 0;
        foreach (string input in SelectInputs(settings, token)) {
            token.ThrowIfCancellationRequested();
            if (++count > settings.MaximumFiles) throw new InvalidDataException("Conversion selection exceeds its file limit.");
            if (SelectRoute(settings, input, routes) is { } route) selected.Add(route.Id);
        }
        return OfficeConverter.Routes.Where(route => selected.Contains(route.Id)).ToArray();
    }

    /// <summary>Replaces the final source extension while preserving the selected relative directory.</summary>
    private static string GetOutputPath(OfficeBatchOptions settings, string input) {
        string relative = settings.InputDirectory == null ? Path.GetFileName(input)
            : OfficeRuntimeCompatibility.GetRelativePath(settings.InputDirectory, input);
        return Path.Combine(settings.OutputDirectory, Path.ChangeExtension(relative, settings.TargetExtension));
    }

    /// <summary>Shares validation and discovery with execution, while never creating output or checkpoint state.</summary>
    internal static Task<OfficeBatchPlan> PlanAsync(IOfficeConversionExecutor runner, OfficeBatchOptions request,
        IProgress<OfficeBatchPlanItem>? progress, CancellationToken token) =>
        Task.Run(() => CreatePlan(runner, request, progress, token), token);

    private static OfficeBatchPlan CreatePlan(IOfficeConversionExecutor runner, OfficeBatchOptions request,
        IProgress<OfficeBatchPlanItem>? progress, CancellationToken token) {
        var result = new OfficeBatchPlan();
        var items = new List<OfficeBatchPlanItem>();
        var notes = new List<string>();
        try {
            token.ThrowIfCancellationRequested();
            OfficeConversionExecutionRoute[] routes = runner.ConversionRoutes.Where(route => route.CanExecute).ToArray();
            OfficeBatchOptions settings = Snapshot(request, routes);
            var destinations = new HashSet<string>(StringComparer.Ordinal);
            if (settings.CheckpointDirectory != null) notes.Add("Checkpoint receipts and output hashes are verified during execution; this plan does not reserve or resume artifacts.");
            foreach (string input in SelectInputs(settings, token)) {
                token.ThrowIfCancellationRequested();
                if (++result.Discovered > settings.MaximumFiles) throw new InvalidDataException("Conversion selection exceeds its file limit.");
                OfficeConversionExecutionRoute? route = SelectRoute(settings, input, routes);
                string? output = null;
                string disposition = "Skipped", message = "No selected executable conversion route.";
                if (route == null) result.Skipped++;
                else {
                    result.Selected++;
                    output = GetOutputPath(settings, input);
                    try {
                        EnsureNoLinks(input); EnsureNoLinks(output);
                        bool directoryPackage = ValidateSelectedInput(runner, settings, route, input);
                        OfficeBatchItemReport? skip = OfficeBatchOutputSelection.ObserveSkip(input, output, settings.SkipExistingOutputs, settings.SkipUpToDate);
                        bool duplicate = !destinations.Add(OfficePathIdentity.NormalizeDirectoryEntry(output));
                        bool collision = duplicate || File.Exists(output) || Directory.Exists(output);
                        if (collision) {
                            result.Collisions++;
                            if (skip == null && settings.ConflictPolicy == OfficePublicationPolicy.Fail && (duplicate || settings.CheckpointDirectory == null))
                                throw new IOException("The proposed output is occupied or shared by selected inputs.");
                        }
                        if (skip != null) {
                            result.Skipped++;
                            disposition = "Skipped";
                            message = skip.SkippedReason + ": " + skip.Summary;
                        } else {
                            disposition = "Selected";
                            message = collision ? "Output collision requires the configured conflict policy or checkpoint verification during execution."
                            : directoryPackage ? "Directory package admitted by its registered owner; member, byte-budget and conversion validation occur during execution."
                            : "Eligible for conversion; document decoding and rendering occur during execution.";
                        }
                    } catch (Exception error) when (error is IOException or InvalidDataException or UnauthorizedAccessException or ArgumentException or NotSupportedException) {
                        result.Failed++; disposition = "Blocked"; message = error.Message;
                    }
                }
                var item = new OfficeBatchPlanItem(input, output, route?.Id, disposition, message);
                if (items.Count < 250) items.Add(item); else result.ItemsTruncated = true;
                progress?.Report(item);
            }
            if (result.Selected == 0) notes.Add("No input files match an executable conversion route.");
        } catch (Exception error) when (error is IOException or InvalidDataException or UnauthorizedAccessException or ArgumentException or NotSupportedException) {
            result.Failed++; notes.Add(error.Message);
        }
        result.Items = items.AsReadOnly();
        result.Diagnostics = notes.Take(32).ToArray();
        return result;
    }

    /// <summary>Keeps preflight and execution admission identical before an output-selection skip.</summary>
    private static bool ValidateSelectedInput(IOfficeConversionExecutor runner, OfficeBatchOptions settings,
        OfficeConversionExecutionRoute route, string input) {
        bool directoryPackage = Directory.Exists(input) && runner.SupportsDirectoryPackageInput(route.Id);
        if (!directoryPackage && !File.Exists(input)) throw new FileNotFoundException("Input file was not found or the configured route does not admit directory packages.");
        if (!directoryPackage && new FileInfo(input).Length > settings.MaximumInputBytes) throw new InvalidDataException("The input exceeds the configured byte limit.");
        if (!route.SupportedOutputProfiles.Contains(settings.OutputProfile))
            throw new ArgumentException("The selected output profile is not supported by this conversion route.");
        _ = settings.ConversionOptions.ForRoute(route);
        if (settings.CheckpointDirectory != null && !route.Checkpointable)
            throw new NotSupportedException("Registered conversion routes require an ordinary batch without checkpoints.");
        return directoryPackage;
    }
}
