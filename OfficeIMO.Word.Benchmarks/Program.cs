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

BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args);
