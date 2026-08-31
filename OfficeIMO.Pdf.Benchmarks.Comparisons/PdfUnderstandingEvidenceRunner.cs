using System.Text.Json;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

internal static class PdfUnderstandingEvidenceRunner {
    internal static int Run(string[] args) {
        string outputPath = ReadOption(args, "--output") ?? throw new ArgumentException("semantic-evidence requires --output <path>.");
        PdfBenchmarkScale scale = ParseScale(ReadOption(args, "--scale") ?? "Medium");
        PdfUnderstandingBenchmarkCorpus corpus = PdfUnderstandingBenchmarkCorpusFactory.Create(scale);
        var pipeline = new PdfUnderstandingPipeline(PdfUnderstandingPipelineOptions.Advanced());
        PdfUnderstandingResult result = pipeline.Run(PdfReadDocument.Open(corpus.Pdf));
        PdfUnderstandingAccuracyObservation accuracy = PdfUnderstandingBenchmarkValidation.Evaluate(result, corpus.Pages);
        PdfLogicalDocument logical = PdfLogicalDocument.Load(corpus.Pdf);
        PdfLogicalStructureObservation logicalStructure = PdfUnderstandingBenchmarkValidation.Observe(logical);
        PdfBinaryClassificationScore tableDetection = PdfUnderstandingBenchmarkValidation.EvaluateTableDetection(logical, corpus.Pages);
        PdfUnderstandingBenchmarkValidation.RequireCompleteLabelCoverage(accuracy);
        PdfUnderstandingBenchmarkValidation.RequireDeterministicSemanticQuality(accuracy);
        PdfUnderstandingBenchmarkValidation.RequireDeterministicTableQuality(tableDetection);

        string fullPath = Path.GetFullPath(outputPath);
        Directory.CreateDirectory(Path.GetDirectoryName(fullPath) ?? Directory.GetCurrentDirectory());
        var report = new {
            schemaVersion = 1,
            generatedAtUtc = DateTimeOffset.UtcNow,
            workload = "OfficeIMO.Pdf.AdvancedUnderstanding",
            scale = scale.ToString(),
            corpus = new {
                source = "deterministic-generated",
                pages = corpus.Pages.Count,
                includes = new[] { "running-header", "spanning-heading", "indented-two-column-reading-order", "list-item", "caption", "table", "running-footer" }
            },
            accuracy,
            tableDetection,
            logicalStructure,
            coverage = new {
                measured = new[] { "labelled-region-character-error-rate", "labelled-region-pairwise-reading-order-accuracy", "labelled-region-kendall-tau", "semantic-kind-precision-recall-f1", "logical-table-detection-precision-recall-f1" },
                notMeasured = new[] { "whole-document-character-error-rate", "heading-level-f1", "cell-adjacency", "teds", "cross-page-continuation-pair-f1", "independent-producer-generalization" }
            },
            interpretation = "This deterministic scorecard is a regression gate, not evidence of accuracy on independent real-world PDFs. Use the separate BenchmarkDotNet workload for elapsed time and managed allocations."
        };
        File.WriteAllText(fullPath, JsonSerializer.Serialize(report, new JsonSerializerOptions { WriteIndented = true }));
        Console.WriteLine(fullPath);
        return 0;
    }

    private static string? ReadOption(string[] args, string name) {
        for (int index = 1; index < args.Length - 1; index++) {
            if (string.Equals(args[index], name, StringComparison.OrdinalIgnoreCase)) {
                return args[index + 1];
            }
        }
        return null;
    }

    private static PdfBenchmarkScale ParseScale(string value) =>
        Enum.TryParse(value, ignoreCase: true, out PdfBenchmarkScale scale)
            ? scale
            : throw new ArgumentException("Unknown semantic-evidence scale: " + value + ". Expected Easy, Medium, or High.");
}
