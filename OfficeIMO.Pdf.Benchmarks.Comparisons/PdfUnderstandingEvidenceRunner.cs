using System.Text.Json;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

internal static class PdfUnderstandingEvidenceRunner {
    internal static int Run(string[] args) {
        string outputPath = ReadOption(args, "--output") ?? throw new ArgumentException("semantic-evidence requires --output <path>.");
        PdfBenchmarkScale scale = ParseScale(ReadOption(args, "--scale") ?? "Medium");
        PdfUnderstandingBenchmarkCorpus corpus = PdfUnderstandingBenchmarkCorpusFactory.Create(scale);
        PdfDocument document = PdfDocument.Load(corpus.Pdf);
        PdfDocumentReadResult structured = document.Read(
            PdfUnderstandingBenchmarkReadOptions.Create(PdfReadProfile.Structured, corpus.Pages.Count));
        PdfSemanticCorrectnessObservation correctness = PdfUnderstandingBenchmarkValidation.Evaluate(structured, corpus);
        PdfStructuredReadObservation structure = PdfUnderstandingBenchmarkValidation.Observe(structured);
        PdfUnderstandingBenchmarkValidation.RequireDeterministicQuality(correctness);

        string fullPath = Path.GetFullPath(outputPath);
        Directory.CreateDirectory(Path.GetDirectoryName(fullPath) ?? Directory.GetCurrentDirectory());
        var report = new {
            schemaVersion = 2,
            generatedAtUtc = DateTimeOffset.UtcNow,
            workload = "OfficeIMO.Pdf.LoadAndRead.Structured",
            scale = scale.ToString(),
            corpus = new {
                source = "deterministic-generated",
                pages = corpus.Pages.Count,
                includes = new[] { "running-header", "spanning-heading", "indented-two-column-reading-order", "list-item", "caption", "table", "running-footer" }
            },
            correctness,
            structure,
            coverage = new {
                measured = new[] {
                    "labelled-region-character-error-rate",
                    "labelled-region-pairwise-reading-order-accuracy",
                    "labelled-region-kendall-tau",
                    "semantic-kind-precision-recall-f1",
                    "heading-detection-precision-recall-f1",
                    "heading-exact-level-precision-recall-f1",
                    "logical-table-detection-precision-recall-f1",
                    "table-cell-adjacency-precision-recall-f1",
                    "cross-page-continuation-pair-precision-recall-f1"
                },
                notMeasured = new[] { "whole-document-character-error-rate", "full-tree-edit-distance-teds", "independent-producer-generalization" }
            },
            interpretation = "This deterministic labelled scorecard is a regression gate, not evidence of generalization to independent real-world PDFs. Cell adjacency is a directly labelled structural score, not a claim of full TEDS equivalence. Use PdfStructuredReadBenchmarks for elapsed time and managed allocations, and the corpus runner for failure rate by document class."
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
