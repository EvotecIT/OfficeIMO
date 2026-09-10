using OfficeIMO.ConversionConsistency;

try {
    string command = args.FirstOrDefault() ?? "help";
    if (command == "help" || command is "--help" or "-h") {
        Console.WriteLine("Conversion consistency: prepare|export|verify|run --suite <json> --output <directory> [--repository <directory>] [--case <id>] [--pdftoppm <executable>]");
        return 0;
    }
    if (command is not ("prepare" or "export" or "verify" or "run"))
        throw new ArgumentException("Unknown command: " + command);
    var seen = new HashSet<string>(StringComparer.Ordinal);
    for (int index = 1; index < args.Length; index += 2) {
        string name = args[index];
        bool allowed = name is "--repository" or "--output" ||
            command is "export" or "run" && name is "--suite" or "--case" ||
            command is "verify" or "run" && name == "--pdftoppm";
        if (!allowed || !seen.Add(name)) throw new ArgumentException("Unknown, duplicate, or inapplicable option: " + name);
        if (index + 1 >= args.Length || args[index + 1].StartsWith("--", StringComparison.Ordinal))
            throw new ArgumentException(name + " requires a value.");
    }
    string repository = Path.GetFullPath(Option("--repository") ?? Environment.CurrentDirectory);
    string output = Path.GetFullPath(Option("--output") ?? throw new ArgumentException("--output is required."));
    if (command == "prepare") { await FixtureCorpus.CreateAsync(repository, output); return 0; }
    using var timeout = new CancellationTokenSource(TimeSpan.FromMinutes(20));
    if (command is "export" or "run") {
        string suitePath = Path.GetFullPath(Option("--suite") ?? throw new ArgumentException("--suite is required."));
        await BundleExporter.ExportAsync(repository, suitePath, output, Option("--case"), timeout.Token);
    }
    if (command is "verify" or "run") {
        GateReport report = await BundleVerifier.VerifyAsync(output, Option("--pdftoppm") ?? "pdftoppm", timeout.Token);
        GateJson.Write(Path.Combine(output, "consistency-result.json"), report);
        Console.WriteLine($"{(report.Passed ? "PASS" : "FAIL")}: {report.Cases.Count} cases; {report.Cases.Sum(item => item.Pages.Count)} pages. Evidence: {output}");
        foreach (CaseReport item in report.Cases.Where(item => !item.Passed)) {
            foreach (string error in item.Errors.Concat(item.Pages.SelectMany(page => page.Errors)))
                Console.Error.WriteLine(item.Id + ": " + error);
        }
        return report.Passed ? 0 : 1;
    }
    if (command != "export") throw new ArgumentException("Unknown command: " + command);
    return 0;
} catch (Exception error) {
    Console.Error.WriteLine(error.Message);
    return 2;
}

string? Option(string name) {
    int index = Array.IndexOf(args, name);
    if (index < 0) return null;
    if (index + 1 >= args.Length || args[index + 1].StartsWith("--", StringComparison.Ordinal))
        throw new ArgumentException(name + " requires a value.");
    return args[index + 1];
}
