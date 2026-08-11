using BenchmarkDotNet.Running;
using OfficeIMO.Word.Benchmarks;

if (args.Length > 0 && string.Equals(args[0], "validate", StringComparison.OrdinalIgnoreCase)) {
    WordLibraryBenchmarkValidation.RunAll();
    Console.WriteLine("All Word library benchmark scenarios produced equivalent validated DOCX payloads.");
    return;
}

if (args.Length > 0 && string.Equals(args[0], "features", StringComparison.OrdinalIgnoreCase)) {
    WordLibraryFeatureCatalog.WriteMarkdown(Console.Out);
    return;
}

BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args);
