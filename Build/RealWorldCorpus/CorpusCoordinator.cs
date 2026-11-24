using System.Runtime.InteropServices;
using OfficeIMO.Reader;

namespace OfficeIMO.RealWorldCorpus;

internal static class CorpusCoordinator {
    public static async Task<CorpusReport> RunAsync(CorpusRunOptions options, CancellationToken cancellationToken = default) {
        DateTimeOffset started = DateTimeOffset.UtcNow;
        // The deadline stops new work, not running workers: each of those still ends within its own timeout, so a
        // run takes at most the deadline plus one per-file timeout and always writes its report.
        using var deadline = new CancellationTokenSource();
        if (options.DeadlineSeconds is int deadlineSeconds) deadline.CancelAfter(TimeSpan.FromSeconds(deadlineSeconds));
        List<CorpusFileRecord> files = Enumerate(options);
        var parallelOptions = new ParallelOptions {
            MaxDegreeOfParallelism = options.Parallelism,
            CancellationToken = cancellationToken
        };

        await Parallel.ForEachAsync(
            files.Where(file => file.Outcome != CorpusOutcomes.SkippedOversize),
            parallelOptions,
            async (file, token) => {
                if (deadline.IsCancellationRequested) {
                    file.Outcome = CorpusOutcomes.SkippedDeadline;
                    return;
                }
                await ClassifyAsync(file, options, token).ConfigureAwait(false);
            }).ConfigureAwait(false);

        Select(files, options);

        await Parallel.ForEachAsync(
            files.Where(file => file.Selected),
            parallelOptions,
            async (file, token) => {
                if (deadline.IsCancellationRequested) {
                    file.Outcome = CorpusOutcomes.SkippedDeadline;
                    return;
                }
                await ProbeAsync(file, options, token).ConfigureAwait(false);
            }).ConfigureAwait(false);

        return BuildReport(options, files, started, DateTimeOffset.UtcNow);
    }

    private static List<CorpusFileRecord> Enumerate(CorpusRunOptions options) {
        var enumeration = new EnumerationOptions {
            RecurseSubdirectories = false,
            IgnoreInaccessible = true,
            AttributesToSkip = 0,
            ReturnSpecialDirectories = false
        };
        var paths = new List<string>();
        var directories = new Stack<string>();
        directories.Push(options.InputDirectory);
        int traversedEntries = 0;
        while (directories.Count > 0) {
            string directory = directories.Pop();
            try {
                foreach (string entry in Directory.EnumerateFileSystemEntries(directory, "*", enumeration)) {
                    if (++traversedEntries > options.MaxTraversalEntries) {
                        throw new InvalidOperationException($"Corpus traversal exceeded --max-traversal-entries ({options.MaxTraversalEntries}).");
                    }

                    FileAttributes attributes;
                    try {
                        attributes = File.GetAttributes(entry);
                    } catch (UnauthorizedAccessException) {
                        continue;
                    } catch (IOException) {
                        continue;
                    }
                    if ((attributes & FileAttributes.ReparsePoint) != 0) continue;
                    if ((attributes & FileAttributes.Directory) != 0) directories.Push(entry);
                    else paths.Add(entry);
                }
            } catch (UnauthorizedAccessException) {
            } catch (IOException) {
            }
        }

        return paths.Select(path => new {
                Path = path,
                Relative = Path.GetRelativePath(options.InputDirectory, path).Replace(Path.DirectorySeparatorChar, '/')
            })
            .OrderBy(item => item.Relative, StringComparer.Ordinal)
            .Select((item, index) => {
                var info = new FileInfo(item.Path);
                return new CorpusFileRecord {
                    FullPath = item.Path,
                    InventoryIndex = index,
                    SourceName = options.IncludeSourceNames ? item.Relative : null,
                    Extension = Path.GetExtension(item.Path).ToLowerInvariant(),
                    SizeBytes = info.Length,
                    Outcome = info.Length > options.MaxFileBytes
                        ? CorpusOutcomes.SkippedOversize
                        : CorpusOutcomes.NotEligible
                };
            })
            .ToList();
    }

    private static async Task ClassifyAsync(CorpusFileRecord file, CorpusRunOptions options, CancellationToken cancellationToken) {
        CorpusProcessResult process = await CorpusProcess.RunAsync(
            "classify-file",
            CorpusOutcomes.Classification,
            file.FullPath,
            options.MaxFileBytes,
            null,
            TimeSpan.FromSeconds(options.TimeoutSeconds),
            cancellationToken).ConfigureAwait(false);
        file.ClassificationDurationMilliseconds = process.DurationMilliseconds;
        if (process.IsTimedOut) {
            file.Outcome = CorpusOutcomes.ClassificationTimedOut;
            file.FailureStage = CorpusOutcomes.Classification;
            return;
        }
        CorpusWorkerResult? worker = process.Worker;
        if (worker == null || !worker.Succeeded) {
            file.Outcome = CorpusOutcomes.ClassificationFailed;
            file.FailureStage = CorpusOutcomes.Classification;
            file.ExceptionType = worker?.ExceptionType ?? process.FailureCode;
            return;
        }
        file.Sha256 = worker.Sha256;
        file.ExtensionKind = worker.ExtensionKind;
        file.ContentKind = worker.ContentKind;
        file.ContentConfidence = worker.ContentConfidence;
        file.DetectedKind = worker.DetectedKind;
        file.Confidence = worker.Confidence;
        file.IsMismatch = worker.IsMismatch;
        file.DetectionEvidence = worker.Evidence;
        if (options.Operation == CorpusOperations.OpenXmlDifferential) {
            // The differential compares how both readers handle what a file claims to be, so an Open XML extension
            // makes it eligible even when its content is damaged or not a package at all.
            ReaderInputKind claimed = CorpusOpenXmlFamilies.FromExtension(file.Extension);
            file.StratumKind = claimed != ReaderInputKind.Unknown
                ? claimed
                : worker.ContentConfidence >= ReaderDetectionConfidence.Medium ? worker.ContentKind : ReaderInputKind.Unknown;
            file.Eligible = options.Formats.Contains(file.StratumKind);
        } else {
            file.StratumKind = worker.ContentKind;
            file.Eligible = options.Formats.Contains(worker.ContentKind) &&
                worker.ContentConfidence >= ReaderDetectionConfidence.Medium;
        }
        file.Outcome = file.Eligible ? CorpusOutcomes.NotSelected : CorpusOutcomes.NotEligible;
    }

    private static void Select(List<CorpusFileRecord> files, CorpusRunOptions options) {
        var seenHashes = new HashSet<string>(StringComparer.Ordinal);
        foreach (CorpusFileRecord file in files
            .Where(file => file.Outcome == CorpusOutcomes.NotSelected && file.Sha256 != null)
            .OrderBy(file => file.Sha256, StringComparer.Ordinal)
            .ThenBy(file => file.InventoryIndex)) {
            if (!seenHashes.Add(file.Sha256!)) file.Outcome = CorpusOutcomes.Duplicate;
        }

        Dictionary<ReaderInputKind, Queue<CorpusFileRecord>> strata = options.Formats.ToDictionary(
            format => format,
            format => new Queue<CorpusFileRecord>(files
                .Where(file => file.Outcome == CorpusOutcomes.NotSelected && file.StratumKind == format)
                .OrderBy(file => file.Sha256, StringComparer.Ordinal)
                .ThenBy(file => file.InventoryIndex)));
        var selectedByFormat = options.Formats.ToDictionary(format => format, _ => 0);
        int selectedTotal = 0;
        bool madeProgress;
        do {
            madeProgress = false;
            foreach (ReaderInputKind format in options.Formats) {
                if (selectedTotal >= options.MaxTotal) return;
                if (selectedByFormat[format] >= options.MaxPerFormat || strata[format].Count == 0) continue;
                CorpusFileRecord file = strata[format].Dequeue();
                file.Selected = true;
                selectedByFormat[format]++;
                selectedTotal++;
                madeProgress = true;
            }
        } while (madeProgress);
    }

    private static async Task ProbeAsync(CorpusFileRecord file, CorpusRunOptions options, CancellationToken cancellationToken) {
        CorpusProcessResult process = await CorpusProcess.RunAsync(
            "probe-file",
            CorpusOutcomes.Probe,
            file.FullPath,
            options.MaxFileBytes,
            file.Sha256,
            TimeSpan.FromSeconds(options.TimeoutSeconds),
            cancellationToken,
            ProbeArguments(options)).ConfigureAwait(false);
        file.ProbeDurationMilliseconds = process.DurationMilliseconds;
        if (process.IsTimedOut) {
            file.Outcome = CorpusOutcomes.TimedOut;
            file.FailureStage = process.TimedOutStage == null ? CorpusOutcomes.Probe : CorpusOutcomes.Probe + ":" + process.TimedOutStage;
            return;
        }
        CorpusWorkerResult? worker = process.Worker;
        if (worker == null || !worker.Succeeded) {
            string? exceptionType = worker?.ExceptionType;
            file.Outcome = IsPolicyRejection(exceptionType)
                ? CorpusOutcomes.Rejected
                : CorpusOutcomes.Failed;
            file.FailureStage = CorpusOutcomes.Probe;
            file.ExceptionType = exceptionType ?? process.FailureCode;
            return;
        }
        file.ChunkCount = worker.ChunkCount;
        file.PageCount = worker.PageCount;
        file.BlockCount = worker.BlockCount;
        file.AssetCount = worker.AssetCount;
        file.InformationDiagnostics = worker.InformationDiagnostics;
        file.WarningDiagnostics = worker.WarningDiagnostics;
        file.ErrorDiagnostics = worker.ErrorDiagnostics;
        file.DiagnosticCodes = worker.DiagnosticCodes;
        file.PdfEvidence = worker.PdfEvidence;
        file.OpenXmlEvidence = worker.OpenXmlEvidence;
        file.Outcome = worker.ErrorDiagnostics > 0
            ? CorpusOutcomes.CompletedWithErrors
            : worker.WarningDiagnostics > 0
                ? CorpusOutcomes.CompletedWithWarnings
                : CorpusOutcomes.Completed;
    }

    private static IReadOnlyList<string>? ProbeArguments(CorpusRunOptions options) {
        if (options.Operation == CorpusOperations.ReaderRead) return null;
        var arguments = new List<string> { "--operation", options.Operation };
        if (options.IncludeDifferenceText) arguments.Add("--include-difference-text");
        return arguments;
    }

    private static CorpusReport BuildReport(
        CorpusRunOptions options,
        List<CorpusFileRecord> files,
        DateTimeOffset started,
        DateTimeOffset completed) {
        var totals = new CorpusTotals {
            Discovered = files.Count,
            Oversize = Count(files, CorpusOutcomes.SkippedOversize),
            ClassificationFailed = Count(files, CorpusOutcomes.ClassificationFailed),
            ClassificationTimedOut = Count(files, CorpusOutcomes.ClassificationTimedOut),
            DuplicateContent = Count(files, CorpusOutcomes.Duplicate),
            EligibleUnique = files.Count(file => file.Eligible && file.Outcome != CorpusOutcomes.Duplicate),
            Selected = files.Count(file => file.Selected),
            Completed = Count(files, CorpusOutcomes.Completed),
            CompletedWithWarnings = Count(files, CorpusOutcomes.CompletedWithWarnings),
            CompletedWithErrors = Count(files, CorpusOutcomes.CompletedWithErrors),
            RejectedByPolicy = Count(files, CorpusOutcomes.Rejected),
            Failed = Count(files, CorpusOutcomes.Failed),
            TimedOut = Count(files, CorpusOutcomes.TimedOut),
            SkippedDeadline = Count(files, CorpusOutcomes.SkippedDeadline),
            PdfSelected = files.Count(file => file.Selected && file.ContentKind == ReaderInputKind.Pdf),
            PdfDeepStages = files.Sum(file => file.PdfEvidence?.Stages.Count ?? 0),
            PdfDeepStagesPassed = files.Sum(file => file.PdfEvidence?.Stages.Count(stage => stage.Succeeded) ?? 0),
            PdfRenderedPages = files.Sum(file => file.PdfEvidence?.RenderSucceededPages ?? 0),
            PdfMutationPlans = files.Sum(file => file.PdfEvidence?.MutationPlanCount ?? 0),
            PdfDeclaredComplianceClaims = files.Sum(file => file.PdfEvidence?.DeclaredComplianceClaimCount ?? 0),
            PdfClaimableComplianceClaims = files.Sum(file => file.PdfEvidence?.ClaimableComplianceClaimCount ?? 0)
        };
        CorpusStratum[] strata = options.Formats.Select(format => {
            CorpusFileRecord[] formatFiles = files.Where(file =>
                file.Eligible && file.StratumKind == format &&
                file.Outcome != CorpusOutcomes.Duplicate).ToArray();
            return new CorpusStratum {
                Format = format,
                EligibleUnique = formatFiles.Length,
                RequestedMaximum = options.MaxPerFormat,
                Selected = formatFiles.Count(file => file.Selected),
                Completed = formatFiles.Count(file => file.Outcome == CorpusOutcomes.Completed),
                CompletedWithWarnings = formatFiles.Count(file => file.Outcome == CorpusOutcomes.CompletedWithWarnings),
                CompletedWithErrors = formatFiles.Count(file => file.Outcome == CorpusOutcomes.CompletedWithErrors),
                RejectedByPolicy = formatFiles.Count(file => file.Outcome == CorpusOutcomes.Rejected),
                Failed = formatFiles.Count(file => file.Outcome == CorpusOutcomes.Failed),
                TimedOut = formatFiles.Count(file => file.Outcome == CorpusOutcomes.TimedOut),
                CorpusUnderfilled = formatFiles.Length < options.MaxPerFormat
            };
        }).ToArray();
        return new CorpusReport {
            StartedUtc = started,
            CompletedUtc = completed,
            Provenance = new CorpusProvenance {
                CorpusId = options.CorpusId,
                SourceUri = options.SourceUri,
                ArchiveSha256 = options.ArchiveSha256,
                Sources = options.SourceProvenancePath == null
                    ? null
                    : System.Text.Json.JsonSerializer.Deserialize<CorpusSourceProvenance[]>(File.ReadAllText(options.SourceProvenancePath), CorpusJson.Options)
            },
            Configuration = new CorpusConfiguration {
                Formats = options.Formats,
                MaxPerFormat = options.MaxPerFormat,
                MaxTotal = options.MaxTotal,
                MaxFileBytes = options.MaxFileBytes,
                MaxTraversalEntries = options.MaxTraversalEntries,
                TimeoutSeconds = options.TimeoutSeconds,
                Parallelism = options.Parallelism,
                DeadlineSeconds = options.DeadlineSeconds,
                SourceNamesIncluded = options.IncludeSourceNames,
                ReaderPolicy = CorpusReaderPolicy.Describe(),
                PackagePolicy = CorpusPackagePolicy.Describe(options.MaxFileBytes),
                Operation = options.Operation == CorpusOperations.OpenXmlDifferential
                    ? "openxml-differential-sdk-reference"
                    : "officeimo-reader-normalized-read"
            },
            Environment = new CorpusEnvironment {
                Framework = RuntimeInformation.FrameworkDescription,
                OperatingSystem = RuntimeInformation.OSDescription,
                ProcessArchitecture = RuntimeInformation.ProcessArchitecture.ToString()
            },
            Totals = totals,
            Strata = strata,
            Files = files.OrderBy(file => file.InventoryIndex).ToArray(),
            OpenXmlLanes = options.Operation == CorpusOperations.OpenXmlDifferential ? SummarizeOpenXmlLanes(files) : null
        };
    }

    internal static CorpusOpenXmlLaneSummary[] SummarizeOpenXmlLanes(IEnumerable<CorpusFileRecord> files) {
        CorpusOpenXmlLane[] lanes = files.Where(file => file.OpenXmlEvidence != null)
            .SelectMany(file => file.OpenXmlEvidence!.Lanes).ToArray();
        return OpenXmlDifferentialProbe.Lanes.Select(name => {
            CorpusOpenXmlLane[] results = lanes.Where(lane => lane.Name == name).ToArray();
            return new CorpusOpenXmlLaneSummary {
                Lane = name,
                Files = results.Length,
                Match = results.Count(lane => lane.Outcome == CorpusOpenXmlOutcomes.Match),
                Mismatch = results.Count(lane => lane.Outcome == CorpusOpenXmlOutcomes.Mismatch),
                OurError = results.Count(lane => lane.Outcome == CorpusOpenXmlOutcomes.OurError),
                SdkError = results.Count(lane => lane.Outcome == CorpusOpenXmlOutcomes.SdkError),
                BothReject = results.Count(lane => lane.Outcome == CorpusOpenXmlOutcomes.BothReject),
                Unavailable = results.Count(lane => lane.Outcome == CorpusOpenXmlOutcomes.Unavailable),
                Streamed = results.Any(lane => lane.Streamed != null) ? results.Count(lane => lane.Streamed == true) : null,
                MismatchClasses = Tally(results.Where(lane => lane.Outcome == CorpusOpenXmlOutcomes.Mismatch).SelectMany(lane => lane.Classes)),
                OurExceptions = Tally(results.Where(lane => lane.Outcome == CorpusOpenXmlOutcomes.OurError).Select(lane => lane.OurException ?? "unknown")),
                StreamedDeclineReasons = Tally(results.Where(lane => lane.StreamedDeclineReason != null).Select(lane => lane.StreamedDeclineReason!))
            };
        }).Where(summary => summary.Files > 0).ToArray();
    }

    private static IReadOnlyDictionary<string, int> Tally(IEnumerable<string> values) =>
        values.GroupBy(value => value, StringComparer.Ordinal)
            .OrderByDescending(group => group.Count()).ThenBy(group => group.Key, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.Count(), StringComparer.Ordinal);

    private static int Count(IEnumerable<CorpusFileRecord> files, string outcome) =>
        files.Count(file => string.Equals(file.Outcome, outcome, StringComparison.Ordinal));

    private static bool IsPolicyRejection(string? exceptionType) =>
        string.Equals(exceptionType, "OfficeIMO.Pdf.PdfReadLimitException", StringComparison.Ordinal) ||
        string.Equals(exceptionType, "OfficeIMO.Pdf.PdfPermissionDeniedException", StringComparison.Ordinal) ||
        string.Equals(exceptionType, "OfficeIMO.Pdf.PdfPasswordRequiredException", StringComparison.Ordinal) ||
        string.Equals(exceptionType, "OfficeIMO.OfficePackageSecurityException", StringComparison.Ordinal) ||
        string.Equals(exceptionType, "OfficeIMO.Html.HtmlDomLimitException", StringComparison.Ordinal) ||
        string.Equals(exceptionType, "OfficeIMO.Rtf.RtfReadLimitException", StringComparison.Ordinal);
}
