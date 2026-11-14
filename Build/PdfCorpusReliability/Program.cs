using System.Globalization;
using OfficeIMO.PdfCorpusReliability;

if (args.Length > 0 && args[0] == "worker") return ReliabilityWorker.Run();
if (args.Length == 0 || args[0] is "-h" or "--help" or "help") {
    RunOptions.PrintHelp();
    return args.Length == 0 ? 1 : 0;
}

try {
    switch (args[0]) {
        case "run": {
            RunOptions options = RunOptions.Parse(args[1..]);
            await ReliabilityCoordinator.RunAsync(options).ConfigureAwait(false);
            return 0;
        }
        case "report": {
            RunOptions options = RunOptions.Parse(args[1..]);
            ReliabilityReport.Write(options);
            return 0;
        }
        default:
            throw new ArgumentException($"Unknown command '{args[0]}'.");
    }
} catch (ArgumentException exception) {
    Console.Error.WriteLine(exception.Message);
    return 2;
}

namespace OfficeIMO.PdfCorpusReliability {
    internal sealed class RunOptions {
        public string CorpusRoot { get; private set; } = "";
        public string? MetadataRoot { get; private set; }
        public string CorpusVersion { get; private set; } = "1.8.1";
        public string OutputDirectory { get; private set; } = "";
        public List<string> Collections { get; } = new();
        public List<string> Only { get; } = new();
        public int? Limit { get; private set; }
        public int Parallelism { get; private set; } = Math.Max(1, Environment.ProcessorCount / 4);
        public TimeSpan Timeout { get; private set; } = TimeSpan.FromSeconds(60);
        public long MemoryLimitBytes { get; private set; } = 2560L * 1024 * 1024;
        public int RenderPages { get; private set; } = 3;
        public bool Resume { get; private set; }
        public string? PngDirectory { get; private set; }

        public string ResultsLinesPath => Path.Combine(OutputDirectory, "results.jsonl");
        public string ResultsJsonPath => Path.Combine(OutputDirectory, "results.json");
        public string SummaryPath => Path.Combine(OutputDirectory, "summary.md");

        public static RunOptions Parse(string[] args) {
            var options = new RunOptions();
            for (int i = 0; i < args.Length; i++) {
                string Next() => i + 1 < args.Length ? args[++i] : throw new ArgumentException($"Missing value for {args[i]}.");
                switch (args[i]) {
                    case "--corpus": options.CorpusRoot = Path.GetFullPath(Next()); break;
                    case "--metadata": options.MetadataRoot = Path.GetFullPath(Next()); break;
                    case "--corpus-version": options.CorpusVersion = Next(); break;
                    case "--output": options.OutputDirectory = Path.GetFullPath(Next()); break;
                    case "--collections": options.Collections.AddRange(Next().Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)); break;
                    case "--only": options.Only.Add(Next()); break;
                    case "--limit": options.Limit = int.Parse(Next(), CultureInfo.InvariantCulture); break;
                    case "--parallel": options.Parallelism = Math.Max(1, int.Parse(Next(), CultureInfo.InvariantCulture)); break;
                    case "--timeout-seconds": options.Timeout = TimeSpan.FromSeconds(int.Parse(Next(), CultureInfo.InvariantCulture)); break;
                    case "--memory-limit-mb": options.MemoryLimitBytes = long.Parse(Next(), CultureInfo.InvariantCulture) * 1024 * 1024; break;
                    case "--render-pages": options.RenderPages = Math.Max(1, int.Parse(Next(), CultureInfo.InvariantCulture)); break;
                    case "--resume": options.Resume = true; break;
                    case "--keep-png": options.PngDirectory = Path.GetFullPath(Next()); break;
                    default: throw new ArgumentException($"Unknown option '{args[i]}'.");
                }
            }
            if (options.CorpusRoot.Length == 0 || !Directory.Exists(options.CorpusRoot)) throw new ArgumentException("--corpus must name an existing directory.");
            if (options.OutputDirectory.Length == 0) throw new ArgumentException("--output is required.");
            if (options.MetadataRoot is not null && !Directory.Exists(options.MetadataRoot)) throw new ArgumentException("--metadata must name an existing directory.");
            if (options.Collections.Count == 0) options.Collections.AddRange(new[] { "regression", "standards", "fuzz" });
            return options;
        }

        public static void PrintHelp() {
            Console.WriteLine("OfficeIMO PDF corpus reliability runner (opt-in; not part of the normal solution)");
            Console.WriteLine();
            Console.WriteLine("  dotnet run -c Release --project Build/PdfCorpusReliability -- run --corpus <dir> --output <dir> [options]");
            Console.WriteLine("  dotnet run -c Release --project Build/PdfCorpusReliability -- report --corpus <dir> --output <dir> [--metadata <dir>]");
            Console.WriteLine();
            Console.WriteLine("Each file is loaded, counted, text-extracted, read structurally, rendered (first/middle/last page, PNG at 72 DPI),");
            Console.WriteLine("saved unmodified, reloaded and compared, and its repair report collected. Files run in recycled child processes");
            Console.WriteLine("with a per-file timeout and working-set cap, so hangs, crashes and memory blow-ups are recorded per file.");
            Console.WriteLine();
            Console.WriteLine("  --metadata <dir>         KillerPDF-Corpus checkout: manifests, published test passwords, qpdf baselines");
            Console.WriteLine("  --corpus-version <v>     manifest version prefix (default 1.8.1)");
            Console.WriteLine("  --collections <list>     regression,standards,fuzz (default all)");
            Console.WriteLine("  --only <path|sha256>     run one file (repeatable); use for repro");
            Console.WriteLine("  --limit <n>              first n files only");
            Console.WriteLine("  --parallel <n>           worker processes (default cores/4)");
            Console.WriteLine("  --timeout-seconds <n>    per file (default 60)");
            Console.WriteLine("  --memory-limit-mb <n>    per worker working set (default 2560)");
            Console.WriteLine("  --render-pages <n>       sampled pages to render (default 3)");
            Console.WriteLine("  --resume                 skip files already in results.jsonl");
            Console.WriteLine("  --keep-png <dir>         write rendered PNGs (repro only)");
            Console.WriteLine();
            Console.WriteLine("Writes results.jsonl (streamed), results.json and summary.md to --output.");
        }
    }
}
