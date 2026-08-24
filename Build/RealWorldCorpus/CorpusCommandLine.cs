using OfficeIMO.Internal;
using OfficeIMO.Reader;

namespace OfficeIMO.RealWorldCorpus;

internal static class CorpusCommandLine {
    private static readonly HashSet<string> RunOptionNames = new(StringComparer.OrdinalIgnoreCase) {
        "--input", "--json", "--markdown", "--corpus-id", "--source-uri", "--archive-sha256",
        "--max-per-format", "--max-total", "--max-file-bytes", "--max-traversal-entries",
        "--timeout-seconds", "--parallelism", "--include-source-names", "--formats"
    };
    private static readonly HashSet<string> WorkerOptionNames = new(StringComparer.OrdinalIgnoreCase) {
        "--stage", "--input", "--max-file-bytes", "--expected-sha256"
    };
    private static readonly HashSet<string> SwitchOptionNames = new(StringComparer.OrdinalIgnoreCase) {
        "--include-source-names"
    };
    private static readonly ReaderInputKind[] DefaultFormats = {
        ReaderInputKind.Word,
        ReaderInputKind.Excel,
        ReaderInputKind.PowerPoint,
        ReaderInputKind.Pdf,
        ReaderInputKind.Html,
        ReaderInputKind.Rtf
    };

    public static CorpusRunOptions ParseRun(string[] args) {
        Dictionary<string, string?> values = Parse(args, RunOptionNames);
        var options = new CorpusRunOptions {
            InputDirectory = Required(values, "--input"),
            JsonReportPath = Required(values, "--json"),
            MarkdownReportPath = Required(values, "--markdown"),
            CorpusId = Optional(values, "--corpus-id") ?? "local-corpus",
            SourceUri = Optional(values, "--source-uri"),
            ArchiveSha256 = Optional(values, "--archive-sha256"),
            MaxPerFormat = PositiveInt(values, "--max-per-format", 100, 1, 10_000),
            MaxTotal = PositiveInt(values, "--max-total", 600, 1, 50_000),
            MaxFileBytes = PositiveLong(values, "--max-file-bytes", 50L * 1024L * 1024L, 256, 1024L * 1024L * 1024L),
            MaxTraversalEntries = PositiveInt(values, "--max-traversal-entries", 5_000, 1, 100_000),
            TimeoutSeconds = PositiveInt(values, "--timeout-seconds", 30, 1, 900),
            Parallelism = PositiveInt(values, "--parallelism", Math.Min(4, Environment.ProcessorCount), 1, 16),
            IncludeSourceNames = values.ContainsKey("--include-source-names"),
            Formats = ParseFormats(Optional(values, "--formats"))
        };

        ResolveAndValidatePaths(options);
        if (options.ArchiveSha256 != null &&
            (options.ArchiveSha256.Length != 64 || options.ArchiveSha256.Any(character => !Uri.IsHexDigit(character)))) {
            throw new ArgumentException("--archive-sha256 must contain exactly 64 hexadecimal characters.");
        }
        return options;
    }

    internal static void ResolveAndValidatePaths(CorpusRunOptions options) {
        options.InputDirectory = OfficePathIdentity.ResolvePhysicalPath(options.InputDirectory);
        options.JsonReportPath = OfficePathIdentity.ResolvePhysicalPath(options.JsonReportPath);
        options.MarkdownReportPath = OfficePathIdentity.ResolvePhysicalPath(options.MarkdownReportPath);
        if (!Directory.Exists(options.InputDirectory)) {
            throw new DirectoryNotFoundException($"Input directory does not exist: {options.InputDirectory}");
        }
        if (string.IsNullOrWhiteSpace(options.CorpusId)) {
            throw new ArgumentException("--corpus-id cannot be empty.");
        }
        if (OfficePathIdentity.AreEquivalent(options.JsonReportPath, options.MarkdownReportPath)) {
            throw new ArgumentException("--json and --markdown must use different output paths.");
        }
        if ((File.Exists(options.JsonReportPath) && OfficePathIdentity.HasMultipleLinks(options.JsonReportPath)) ||
            (File.Exists(options.MarkdownReportPath) && OfficePathIdentity.HasMultipleLinks(options.MarkdownReportPath))) {
            throw new ArgumentException("Existing report files must not have multiple hard links.");
        }
        if (OfficePathIdentity.IsSameOrDescendant(options.JsonReportPath, options.InputDirectory) ||
            OfficePathIdentity.IsSameOrDescendant(options.MarkdownReportPath, options.InputDirectory)) {
            throw new ArgumentException("Report paths must be outside --input so prior evidence cannot enter a later sample.");
        }
    }

    public static CorpusWorkerOptions ParseWorker(string[] args) {
        Dictionary<string, string?> values = Parse(args, WorkerOptionNames);
        string stage = Required(values, "--stage");
        if (stage is not (CorpusOutcomes.Classification or CorpusOutcomes.Probe)) {
            throw new ArgumentException("--stage must be 'classification' or 'probe'.");
        }
        string? expectedSha256 = Optional(values, "--expected-sha256");
        if (stage == CorpusOutcomes.Probe) {
            ValidateSha256(expectedSha256, "--expected-sha256");
        } else if (expectedSha256 != null) {
            throw new ArgumentException("--expected-sha256 is valid only for the probe stage.");
        }
        return new CorpusWorkerOptions {
            InputPath = Path.GetFullPath(Required(values, "--input")),
            MaxFileBytes = PositiveLong(values, "--max-file-bytes", 50L * 1024L * 1024L, 256, 1024L * 1024L * 1024L),
            ExpectedSha256 = expectedSha256?.ToLowerInvariant(),
            Stage = stage
        };
    }

    private static Dictionary<string, string?> Parse(string[] args, IReadOnlySet<string> allowedNames) {
        var values = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
        for (int index = 0; index < args.Length; index++) {
            string name = args[index];
            if (!name.StartsWith("--", StringComparison.Ordinal)) {
                throw new ArgumentException($"Unexpected argument '{name}'.");
            }
            if (!allowedNames.Contains(name)) {
                throw new ArgumentException($"Unknown option '{name}'.");
            }
            if (values.ContainsKey(name)) {
                throw new ArgumentException($"Option '{name}' was specified more than once.");
            }
            if (SwitchOptionNames.Contains(name)) {
                values.Add(name, null);
                continue;
            }
            if (index + 1 >= args.Length || args[index + 1].StartsWith("--", StringComparison.Ordinal)) {
                throw new ArgumentException($"{name} requires a value.");
            }
            values.Add(name, args[++index]);
        }
        return values;
    }

    private static string Required(IReadOnlyDictionary<string, string?> values, string name) =>
        Optional(values, name) ?? throw new ArgumentException($"{name} is required.");

    private static string? Optional(IReadOnlyDictionary<string, string?> values, string name) =>
        values.TryGetValue(name, out string? value) ? value : null;

    private static int PositiveInt(IReadOnlyDictionary<string, string?> values, string name, int defaultValue, int minimum, int maximum) {
        string? text = Optional(values, name);
        if (text == null) return defaultValue;
        if (!int.TryParse(text, out int value) || value < minimum || value > maximum) {
            throw new ArgumentOutOfRangeException(name, $"Expected an integer from {minimum} through {maximum}.");
        }
        return value;
    }

    private static long PositiveLong(IReadOnlyDictionary<string, string?> values, string name, long defaultValue, long minimum, long maximum) {
        string? text = Optional(values, name);
        if (text == null) return defaultValue;
        if (!long.TryParse(text, out long value) || value < minimum || value > maximum) {
            throw new ArgumentOutOfRangeException(name, $"Expected an integer from {minimum} through {maximum}.");
        }
        return value;
    }

    private static IReadOnlyList<ReaderInputKind> ParseFormats(string? value) {
        if (string.IsNullOrWhiteSpace(value)) return DefaultFormats;
        ReaderInputKind[] formats = value.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Select(item => Enum.TryParse(item, ignoreCase: true, out ReaderInputKind format)
                ? format
                : throw new ArgumentException($"Unknown format '{item}' in --formats."))
            .Distinct()
            .ToArray();
        if (formats.Length == 0 || formats.Any(format => !DefaultFormats.Contains(format))) {
            throw new ArgumentException("--formats supports Word, Excel, PowerPoint, Pdf, Html, and Rtf.");
        }
        return formats;
    }

    private static void ValidateSha256(string? value, string optionName) {
        if (value == null || value.Length != 64 || value.Any(character => !Uri.IsHexDigit(character))) {
            throw new ArgumentException($"{optionName} must contain exactly 64 hexadecimal characters.");
        }
    }
}
