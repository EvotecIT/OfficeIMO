using System.Text.Json;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

internal static class PdfCompatibilityRunner {
    internal static int Run(string[] args) {
        string outputDirectory = args.Length > 1
            ? Path.GetFullPath(args[1])
            : Path.GetFullPath(Path.Combine("Ignore", "Benchmarks", "PdfComparisons", "compatibility"));
        Directory.CreateDirectory(outputDirectory);

        var rows = new List<PdfCompatibilityRow>();
        foreach (PdfBenchmarkScale scale in Enum.GetValues<PdfBenchmarkScale>()) {
            PdfBenchmarkScenario scenario = PdfBenchmarkScenario.Get(scale);
            foreach (PdfBenchmarkProducer producer in Enum.GetValues<PdfBenchmarkProducer>()) {
                byte[] pdf = PdfDocumentGenerators.Generate(producer, scenario);
                PdfBenchmarkValidation.ValidateGenerated(pdf, scenario, producer.ToString());
                if (scale == PdfBenchmarkScale.Easy) {
                    File.WriteAllBytes(Path.Combine(outputDirectory, producer + ".pdf"), pdf);
                    File.WriteAllText(
                        Path.Combine(outputDirectory, producer + ".debug.txt"),
                        OfficeIMO.Pdf.PdfDocument.Open(pdf).Debug(new OfficeIMO.Pdf.PdfDebuggerOptions {
                            IncludeDecodedStreamPreviews = true,
                            MaxDecodedStreamPreviewBytes = 32 * 1024
                        }).ToText());
                }

                foreach (PdfReaderEngine reader in Enum.GetValues<PdfReaderEngine>()) {
                    PdfReadObservation observation = default;
                    try {
                        observation = PdfDocumentReaders.Read(reader, pdf);
                        if (scale == PdfBenchmarkScale.Easy) {
                            File.WriteAllText(
                                Path.Combine(outputDirectory, producer + "." + reader + ".text.txt"),
                                PdfDocumentReaders.ExtractText(reader, pdf));
                        }
                        PdfBenchmarkValidation.ValidateRead(observation, scenario, reader + " reading " + producer);
                        rows.Add(new PdfCompatibilityRow(scale, producer, reader, true, observation, null));
                    } catch (Exception exception) {
                        rows.Add(new PdfCompatibilityRow(scale, producer, reader, false, observation, exception.Message));
                    }
                }
            }
        }

        string reportPath = Path.Combine(outputDirectory, "compatibility.json");
        File.WriteAllText(reportPath, JsonSerializer.Serialize(rows, new JsonSerializerOptions { WriteIndented = true }));
        foreach (PdfCompatibilityRow row in rows) {
            Console.WriteLine(
                $"{row.Scale,-6} {row.Producer,-9} {row.Reader,-9} {(row.Success ? "PASS" : "FAIL")} " +
                $"pages={row.Observation.PageCount} markers={row.Observation.ReportMarkerCount} text={row.Observation.TextLength}" +
                (row.Error == null ? string.Empty : " error=" + row.Error));
        }

        Console.WriteLine($"Compatibility report: {reportPath}");
        return rows.All(static row => row.Success) ? 0 : 2;
    }

    private sealed record PdfCompatibilityRow(
        PdfBenchmarkScale Scale,
        PdfBenchmarkProducer Producer,
        PdfReaderEngine Reader,
        bool Success,
        PdfReadObservation Observation,
        string? Error);
}
