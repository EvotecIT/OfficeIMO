using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Exporters.Json;
using BenchmarkDotNet.Running;
using OfficeIMO.Pdf.Benchmarks.Comparisons;

if (args.Length > 0 && string.Equals(args[0], "compatibility", StringComparison.OrdinalIgnoreCase)) {
    return PdfCompatibilityRunner.Run(args);
}

if (args.Length > 0 && string.Equals(args[0], "corpus", StringComparison.OrdinalIgnoreCase)) {
    return await PdfCorpusRunner.RunAsync(args);
}

if (args.Length > 0 && string.Equals(args[0], "semantic-evidence", StringComparison.OrdinalIgnoreCase)) {
    return PdfUnderstandingEvidenceRunner.Run(args);
}

if (args.Length > 0 && string.Equals(args[0], "html-package-evidence", StringComparison.OrdinalIgnoreCase)) {
    string outputPath = Path.GetFullPath(ReadOption(args, "--output")
        ?? throw new ArgumentException("html-package-evidence requires --output."));
    Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);
    File.WriteAllText(outputPath, System.Text.Json.JsonSerializer.Serialize(PdfPeckerPackageEvidence.Read(),
        new System.Text.Json.JsonSerializerOptions { PropertyNamingPolicy = System.Text.Json.JsonNamingPolicy.CamelCase, WriteIndented = true }));
    return 0;
}

if (args.Length > 0 && string.Equals(args[0], "html-payload-evidence", StringComparison.OrdinalIgnoreCase)) {
    return PdfHtmlPayloadEvidenceRunner.Run(args);
}

if (args.Length > 0 && string.Equals(args[0], "html-evidence", StringComparison.OrdinalIgnoreCase)) {
    return await HtmlPdfEvidenceRunner.RunAsync(args);
}

if (args.Length > 0 && string.Equals(args[0], "html-mhtml-evidence", StringComparison.OrdinalIgnoreCase)) {
    return await HtmlMhtmlEvidenceRunner.RunAsync(args);
}

if (args.Length > 0 && string.Equals(args[0], "html-evidence-worker", StringComparison.OrdinalIgnoreCase)) {
    return await HtmlPdfEvidenceWorker.RunAsync(args);
}

if (args.Length > 0 && string.Equals(args[0], "invoice-evidence", StringComparison.OrdinalIgnoreCase)) {
    return await PdfInvoiceEvidenceRunner.RunAsync(args);
}

if (args.Length > 0 && string.Equals(args[0], "prepare-rich-word", StringComparison.OrdinalIgnoreCase)) {
    string repositoryRoot = ReadOption(args, "--repo-root") ?? FindRepositoryRoot();
    string outputDirectory = Path.GetFullPath(
        ReadOption(args, "--output") ??
        Path.Combine(repositoryRoot, "Ignore", "Benchmarks", "PdfComparisons", "rich-word"));
    RichWordPdfCorpusArtifacts artifacts = RichWordPdfCorpusGenerator.Generate(repositoryRoot, outputDirectory);
    Console.WriteLine("DOCX=" + artifacts.DocxPath);
    Console.WriteLine("PDF=" + artifacts.PdfPath);
    Console.WriteLine("CONVERSION_REPORT=" + artifacts.ConversionReportPath);
    return 0;
}

ManualConfig config = ManualConfig
    .Create(DefaultConfig.Instance)
    .AddExporter(JsonExporter.Full);

BenchmarkSwitcher
    .FromAssembly(typeof(PdfGenerationBenchmarks).Assembly)
    .Run(args, config);

return 0;

static string? ReadOption(string[] values, string option) {
    for (int index = 1; index < values.Length - 1; index++) {
        if (string.Equals(values[index], option, StringComparison.OrdinalIgnoreCase)) {
            return values[index + 1];
        }
    }

    return null;
}

static string FindRepositoryRoot() {
    string? current = Directory.GetCurrentDirectory();
    while (!string.IsNullOrWhiteSpace(current)) {
        if (File.Exists(Path.Combine(current, "OfficeIMO.sln"))) {
            return current;
        }

        current = Directory.GetParent(current)?.FullName;
    }

    throw new DirectoryNotFoundException("Could not locate the OfficeIMO repository root.");
}
