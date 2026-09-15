using BenchmarkDotNet.Running;
using OfficeIMO.Word.Benchmarks;

if (args.Length > 0 && string.Equals(args[0], "--evidence-probe", StringComparison.OrdinalIgnoreCase)) {
    Environment.ExitCode = WordOpenXmlEvidenceRunner.RunProbe(args[1..]);
    return;
}

if (args.Length > 0 && string.Equals(args[0], "evidence", StringComparison.OrdinalIgnoreCase)) {
    Environment.ExitCode = WordOpenXmlEvidenceRunner.Run(args[1..]);
    return;
}

if (args.Length > 0 && string.Equals(args[0], "validate-openxml", StringComparison.OrdinalIgnoreCase)) {
    WordOpenXmlEvidenceValidation.RunAll();
    Console.WriteLine("OfficeIMO and Open XML SDK produced equivalent validated DOCX payloads.");
    return;
}

if (args.Length > 0 && string.Equals(args[0], "validate", StringComparison.OrdinalIgnoreCase)) {
    WordLibraryBenchmarkValidation.RunAll();
    Console.WriteLine("All Word library benchmark scenarios produced equivalent validated DOCX payloads.");
    return;
}

if (args.Length > 0 && string.Equals(args[0], "validate-workflows", StringComparison.OrdinalIgnoreCase)) {
    foreach (int itemCount in new[] { 100, 1000 }) {
        var workload = new WordWorkflowBenchmarks { ItemCount = itemCount };
        try {
            workload.Setup();
        } finally {
            workload.Cleanup();
        }
    }
    Console.WriteLine("Word workflow benchmark setup validated the 100- and 1000-item semantic contracts.");
    return;
}

BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args);
