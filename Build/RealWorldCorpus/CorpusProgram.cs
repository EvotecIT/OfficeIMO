using System.Text.Json;

namespace OfficeIMO.RealWorldCorpus;

internal static class CorpusProgram {
    public static async Task<int> RunAsync(string[] args) {
        if (args.Length == 0 || IsHelp(args[0])) {
            PrintHelp();
            return args.Length == 0 ? 1 : 0;
        }

        try {
            return args[0].ToLowerInvariant() switch {
                "run" => await RunCorpusAsync(CorpusCommandLine.ParseRun(args[1..])).ConfigureAwait(false),
                "fetch-sources" => await CorpusSources.FetchAsync(args[1..]).ConfigureAwait(false),
                "verify-markdown-contract" => CorpusMarkdownContract.Run(),
                "classify-file" => RunWorker(args[1..]),
                "probe-file" => RunWorker(args[1..]),
                _ => throw new ArgumentException($"Unknown command '{args[0]}'.")
            };
        } catch (Exception exception) {
            Console.Error.WriteLine($"{exception.GetType().Name}: {exception.Message}");
            return 1;
        }
    }

    private static async Task<int> RunCorpusAsync(CorpusRunOptions options) {
        CorpusReport report = await CorpusCoordinator.RunAsync(options).ConfigureAwait(false);
        CorpusCommandLine.ResolveAndValidatePaths(options);
        CorpusReportWriter.Write(report, options.JsonReportPath, options.MarkdownReportPath);
        Console.WriteLine($"Measured {report.Totals.Selected} selected files from {report.Totals.Discovered} discovered files.");
        foreach (CorpusOpenXmlLaneSummary lane in report.OpenXmlLanes ?? Array.Empty<CorpusOpenXmlLaneSummary>()) {
            Console.WriteLine($"{lane.Lane}: {lane.Match}/{lane.Compared} match, {lane.Mismatch} mismatch, {lane.OurError} our-error, {lane.SdkError} sdk-error, {lane.BothReject} both-reject, {lane.Unavailable} unavailable");
        }
        Console.WriteLine($"JSON: {options.JsonReportPath}");
        Console.WriteLine($"Markdown: {options.MarkdownReportPath}");
        return 0;
    }

    private static int RunWorker(string[] args) {
        CorpusWorkerOptions options = CorpusCommandLine.ParseWorker(args);
        CorpusWorkerResult result;
        try {
            result = options.Stage == CorpusOutcomes.Classification
                ? CorpusWorker.Classify(options.InputPath, options.MaxFileBytes)
                : options.Operation == CorpusOperations.OpenXmlDifferential
                    ? CorpusWorker.ProbeOpenXml(options.InputPath, options.MaxFileBytes, options.ExpectedSha256!, options.IncludeDifferenceText)
                    : CorpusWorker.Probe(options.InputPath, options.MaxFileBytes, options.ExpectedSha256!);
        } catch (Exception exception) {
            result = CorpusWorkerResult.Failure(options.Stage, exception);
        }

        Console.Out.Write(JsonSerializer.Serialize(result, CorpusJson.Options));
        return 0;
    }

    private static bool IsHelp(string value) => value is "-h" or "--help" or "help";

    private static void PrintHelp() {
        Console.WriteLine("OfficeIMO real-world corpus evidence runner");
        Console.WriteLine();
        Console.WriteLine("Usage:");
        Console.WriteLine("  dotnet run --project Build/RealWorldCorpus/OfficeIMO.RealWorldCorpus.Tool.csproj --framework net10.0 -- run \\");
        Console.WriteLine("    --input <directory> --json <report.json> --markdown <report.md> [--operation reader|openxml-differential] [options]");
        Console.WriteLine("  dotnet run --project Build/RealWorldCorpus/OfficeIMO.RealWorldCorpus.Tool.csproj --framework net10.0 -- fetch-sources \\");
        Console.WriteLine("    --manifest Build/RealWorldCorpus/openxml-corpus-sources.json --output <directory> [--sources id,...] [--print-pins]");
        Console.WriteLine();
        Console.WriteLine("The runner content-detects files, selects a deterministic hash-ordered sample per format,");
        Console.WriteLine("and probes each selected file in a separately timed process. See the evidence guide for limits.");
    }
}
