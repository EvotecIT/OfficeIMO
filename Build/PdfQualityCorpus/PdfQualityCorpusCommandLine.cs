namespace OfficeIMO.PdfQualityCorpus;

internal static class PdfQualityCorpusCommandLine {
    internal static QualityRunOptions ParseRun(string[] args) {
        Dictionary<string, string> values = ParseValues(args, 1);
        ValidateNames(values, "--manifest", "--root", "--json", "--markdown", "--max-file-bytes", "--max-render-pages", "--timeout-seconds", "--parallelism", "--max-worker-memory-bytes");
        var options = new QualityRunOptions {
            ManifestPath = Required(values, "--manifest"),
            RootDirectory = Required(values, "--root"),
            JsonReportPath = Required(values, "--json"),
            MarkdownReportPath = Required(values, "--markdown"),
            MaxFileBytes = Long(values, "--max-file-bytes", 128L * 1024L * 1024L),
            MaxRenderPages = Integer(values, "--max-render-pages", 4),
            TimeoutSeconds = Integer(values, "--timeout-seconds", 60),
            Parallelism = Integer(values, "--parallelism", Math.Max(1, Math.Min(4, Environment.ProcessorCount))),
            MaxWorkerMemoryBytes = Long(values, "--max-worker-memory-bytes", 1024L * 1024L * 1024L)
        };
        ValidatePositive(options.MaxFileBytes, "--max-file-bytes");
        ValidatePositive(options.MaxRenderPages, "--max-render-pages");
        ValidatePositive(options.TimeoutSeconds, "--timeout-seconds");
        ValidatePositive(options.Parallelism, "--parallelism");
        ValidatePositive(options.MaxWorkerMemoryBytes, "--max-worker-memory-bytes");
        options.ManifestPath = Path.GetFullPath(options.ManifestPath);
        options.RootDirectory = Path.GetFullPath(options.RootDirectory);
        options.JsonReportPath = Path.GetFullPath(options.JsonReportPath);
        options.MarkdownReportPath = Path.GetFullPath(options.MarkdownReportPath);
        if (!File.Exists(options.ManifestPath)) throw new FileNotFoundException("PDF quality corpus manifest was not found.", options.ManifestPath);
        if (!Directory.Exists(options.RootDirectory)) throw new DirectoryNotFoundException("PDF quality corpus root was not found: " + options.RootDirectory);
        if (string.Equals(options.JsonReportPath, options.MarkdownReportPath, StringComparison.OrdinalIgnoreCase)) {
            throw new ArgumentException("JSON and Markdown report paths must be different.");
        }
        return options;
    }

    internal static QualityProbeOptions ParseProbe(string[] args) {
        Dictionary<string, string> values = ParseValues(args, 1);
        ValidateNames(values, "--manifest", "--root", "--case-id", "--max-file-bytes", "--max-render-pages");
        var options = new QualityProbeOptions {
            ManifestPath = Path.GetFullPath(Required(values, "--manifest")),
            RootDirectory = Path.GetFullPath(Required(values, "--root")),
            CaseId = Required(values, "--case-id"),
            MaxFileBytes = Long(values, "--max-file-bytes", 128L * 1024L * 1024L),
            MaxRenderPages = Integer(values, "--max-render-pages", 4)
        };
        ValidatePositive(options.MaxFileBytes, "--max-file-bytes");
        ValidatePositive(options.MaxRenderPages, "--max-render-pages");
        return options;
    }

    private static void ValidateNames(Dictionary<string, string> values, params string[] allowedNames) {
        var allowed = new HashSet<string>(allowedNames, StringComparer.Ordinal);
        string? unknown = values.Keys.FirstOrDefault(name => !allowed.Contains(name));
        if (unknown is not null) throw new ArgumentException("Unknown option: " + unknown + ".");
    }

    private static Dictionary<string, string> ParseValues(string[] args, int startIndex) {
        var values = new Dictionary<string, string>(StringComparer.Ordinal);
        for (int i = startIndex; i < args.Length; i += 2) {
            if (i + 1 >= args.Length || !args[i].StartsWith("--", StringComparison.Ordinal)) {
                throw new ArgumentException("Options must use --name value pairs.");
            }
            if (!values.TryAdd(args[i], args[i + 1])) throw new ArgumentException("Duplicate option: " + args[i] + ".");
        }
        return values;
    }

    private static string Required(Dictionary<string, string> values, string name) =>
        values.TryGetValue(name, out string? value) && !string.IsNullOrWhiteSpace(value)
            ? value
            : throw new ArgumentException("Missing required option " + name + ".");

    private static int Integer(Dictionary<string, string> values, string name, int fallback) =>
        values.TryGetValue(name, out string? value) && int.TryParse(value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int parsed)
            ? parsed
            : values.ContainsKey(name) ? throw new ArgumentException("Invalid integer for " + name + ".") : fallback;

    private static long Long(Dictionary<string, string> values, string name, long fallback) =>
        values.TryGetValue(name, out string? value) && long.TryParse(value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out long parsed)
            ? parsed
            : values.ContainsKey(name) ? throw new ArgumentException("Invalid integer for " + name + ".") : fallback;

    private static void ValidatePositive(long value, string name) {
        if (value <= 0) throw new ArgumentOutOfRangeException(name, value, "Value must be positive.");
    }
}
