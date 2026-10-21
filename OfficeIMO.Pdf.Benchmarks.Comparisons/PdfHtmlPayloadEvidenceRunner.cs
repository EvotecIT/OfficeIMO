using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

/// <summary>Exercises the exact benchmark payloads without measuring or ranking their execution.</summary>
internal static class PdfHtmlPayloadEvidenceRunner {
    internal static int Run(string[] args) {
        if (args.Length != 3 || args[1] != "--output") {
            throw new ArgumentException("html-payload-evidence requires --output <directory>.");
        }
        string outputDirectory = Path.GetFullPath(args[2]);
        using EvidenceOutputReservation reservation = EvidenceOutputReservation.Acquire(outputDirectory);
        var cases = new List<PayloadEvidence>();
        foreach (PdfHtmlPayloadKind payload in Enum.GetValues<PdfHtmlPayloadKind>()) {
            PdfHtmlPayloadScenario scenario = PdfHtmlPayloadScenario.Create(payload);
            byte[] input = Encoding.UTF8.GetBytes(scenario.Html);
            string inputName = payload + ".html";
            File.WriteAllBytes(Path.Combine(outputDirectory, inputName), input);
            var benchmark = new PdfHtmlPayloadBenchmarks { Payload = payload };
            benchmark.Setup();
            foreach (HtmlPdfComparisonEngine engine in new[] {
                HtmlPdfComparisonEngine.OfficeIMO, HtmlPdfComparisonEngine.PeachPDF, HtmlPdfComparisonEngine.PdfPecker }) {
                string fileName = payload + "-" + engine + ".pdf";
                byte[]? bytes = null;
                string? failure = null;
                PdfReadObservation? observation = null;
                try {
                    bytes = engine switch {
                        HtmlPdfComparisonEngine.OfficeIMO => benchmark.OfficeIMO(),
                        HtmlPdfComparisonEngine.PeachPDF => benchmark.PeachPDF(),
                        HtmlPdfComparisonEngine.PdfPecker => benchmark.PdfPecker(),
                        _ => throw new ArgumentOutOfRangeException(nameof(engine))
                    };
                    File.WriteAllBytes(Path.Combine(outputDirectory, fileName), bytes);
                    observation = benchmark.Validate(engine.ToString(), bytes);
                } catch (Exception exception) {
                    failure = exception.GetType().Name + ": " + exception.Message;
                }
                cases.Add(new(payload.ToString(), engine.ToString(), failure == null ? "Passed" : "Rejected",
                    inputName, input.Length, Hash(input), bytes == null ? null : fileName,
                    bytes?.LongLength, bytes == null ? null : Hash(bytes), observation?.PageCount, observation?.TextLength, failure));
            }
        }
        bool accepted = cases.All(item => item.Status == "Passed");
        string reportPath = Path.Combine(outputDirectory, "html-payload-evidence.json");
        File.WriteAllText(reportPath, JsonSerializer.Serialize(new {
            schemaVersion = 1,
            generatedUtc = DateTimeOffset.UtcNow,
            measurementScope = "Correctness only; no timing or speed ranking",
            source = SourceProvenanceReader.ReadGitStateAsync(Directory.GetCurrentDirectory()).GetAwaiter().GetResult(),
            runtimeVersion = Environment.Version.ToString(),
            osDescription = System.Runtime.InteropServices.RuntimeInformation.OSDescription,
            fontProfile = new {
                name = "SharedCarlitoRegularBold",
                regularSha256 = Hash(PdfBenchmarkAssets.CarlitoRegular),
                boldSha256 = Hash(PdfBenchmarkAssets.CarlitoBold)
            },
            package = PdfPeckerPackageEvidence.Read(),
            allCasesAccepted = accepted,
            cases
        }, new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase, WriteIndented = true }));
        Console.WriteLine("HTML_PAYLOAD_EVIDENCE_REPORT=" + reportPath);
        return accepted ? 0 : 1;
    }

    private static string Hash(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

    private sealed record PayloadEvidence(string Payload, string Engine, string Status,
        string InputRelativePath, int InputBytes, string InputSha256, string? PdfRelativePath,
        long? PdfBytes, string? PdfSha256, int? PageCount, int? TextLength, string? Failure);
}
