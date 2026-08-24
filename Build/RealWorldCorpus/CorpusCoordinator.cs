using System.Runtime.InteropServices;
using OfficeIMO.Reader;

namespace OfficeIMO.RealWorldCorpus;

internal static class CorpusCoordinator {
    public static async Task<CorpusReport> RunAsync(CorpusRunOptions options, CancellationToken cancellationToken = default) {
        DateTimeOffset started = DateTimeOffset.UtcNow;
        List<CorpusFileRecord> files = Enumerate(options);
        var parallelOptions = new ParallelOptions {
            MaxDegreeOfParallelism = options.Parallelism,
            CancellationToken = cancellationToken
        };

        await Parallel.ForEachAsync(
            files.Where(file => file.Outcome != CorpusOutcomes.SkippedOversize),
            parallelOptions,
            async (file, token) => await ClassifyAsync(file, options, token).ConfigureAwait(false)).ConfigureAwait(false);

        Select(files, options);

        await Parallel.ForEachAsync(
            files.Where(file => file.Selected),
            parallelOptions,
            async (file, token) => await ProbeAsync(file, options, token).ConfigureAwait(false)).ConfigureAwait(false);

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
        file.Outcome = options.Formats.Contains(worker.ContentKind) &&
            worker.ContentConfidence >= ReaderDetectionConfidence.Medium
            ? CorpusOutcomes.NotSelected
            : CorpusOutcomes.NotEligible;
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
                .Where(file => file.Outcome == CorpusOutcomes.NotSelected && file.ContentKind == format)
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
            cancellationToken).ConfigureAwait(false);
        file.ProbeDurationMilliseconds = process.DurationMilliseconds;
        if (process.IsTimedOut) {
            file.Outcome = CorpusOutcomes.TimedOut;
            file.FailureStage = CorpusOutcomes.Probe;
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
        file.Outcome = worker.ErrorDiagnostics > 0
            ? CorpusOutcomes.CompletedWithErrors
            : worker.WarningDiagnostics > 0
                ? CorpusOutcomes.CompletedWithWarnings
                : CorpusOutcomes.Completed;
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
            EligibleUnique = files.Count(file =>
                options.Formats.Contains(file.ContentKind) &&
                file.ContentConfidence >= ReaderDetectionConfidence.Medium &&
                file.Outcome != CorpusOutcomes.Duplicate),
            Selected = files.Count(file => file.Selected),
            Completed = Count(files, CorpusOutcomes.Completed),
            CompletedWithWarnings = Count(files, CorpusOutcomes.CompletedWithWarnings),
            CompletedWithErrors = Count(files, CorpusOutcomes.CompletedWithErrors),
            RejectedByPolicy = Count(files, CorpusOutcomes.Rejected),
            Failed = Count(files, CorpusOutcomes.Failed),
            TimedOut = Count(files, CorpusOutcomes.TimedOut)
        };
        CorpusStratum[] strata = options.Formats.Select(format => {
            CorpusFileRecord[] formatFiles = files.Where(file =>
                file.ContentKind == format &&
                file.ContentConfidence >= ReaderDetectionConfidence.Medium &&
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
                ArchiveSha256 = options.ArchiveSha256
            },
            Configuration = new CorpusConfiguration {
                Formats = options.Formats,
                MaxPerFormat = options.MaxPerFormat,
                MaxTotal = options.MaxTotal,
                MaxFileBytes = options.MaxFileBytes,
                MaxTraversalEntries = options.MaxTraversalEntries,
                TimeoutSeconds = options.TimeoutSeconds,
                Parallelism = options.Parallelism,
                SourceNamesIncluded = options.IncludeSourceNames,
                PackagePolicy = CorpusPackagePolicy.Describe(options.MaxFileBytes)
            },
            Environment = new CorpusEnvironment {
                Framework = RuntimeInformation.FrameworkDescription,
                OperatingSystem = RuntimeInformation.OSDescription,
                ProcessArchitecture = RuntimeInformation.ProcessArchitecture.ToString()
            },
            Totals = totals,
            Strata = strata,
            Files = files.OrderBy(file => file.InventoryIndex).ToArray()
        };
    }

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
