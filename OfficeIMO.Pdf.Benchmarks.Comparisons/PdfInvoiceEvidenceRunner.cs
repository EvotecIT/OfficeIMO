using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text.Json;
using OfficeIMO.Pdf;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

internal static class PdfInvoiceEvidenceRunner {
    internal static async Task<int> RunAsync(string[] args) {
        string repositoryRoot = FindRepositoryRoot();
        GitSourceState provenance = await SourceProvenanceReader.ReadGitStateAsync(repositoryRoot).ConfigureAwait(false)
            ?? throw new InvalidOperationException("Could not read commit-addressable OfficeIMO source provenance.");
        if (HasOption(args, "--require-clean-source") && !provenance.IsClean) {
            throw new InvalidOperationException("--require-clean-source was specified, but the OfficeIMO worktree has changes.");
        }

        string output = Path.GetFullPath(ReadOption(args, "--output") ??
            Path.Combine("Ignore", "Benchmarks", "PdfInvoiceEvidence"));
        Directory.CreateDirectory(output);

        InvoiceComparisonScenario scenario = InvoiceComparisonScenario.Create();
        var artifacts = new[] {
            Write(output, "officeimo", OfficeImoPdfInvoiceGenerator.Generate(scenario), scenario),
            Write(output, "questpdf", QuestPdfInvoiceGenerator.Generate(scenario), scenario),
            Write(output, "itext", ITextPdfInvoiceGenerator.Generate(scenario), scenario)
        };

        string reportPath = Path.Combine(output, "invoice-evidence.json");
        File.WriteAllText(reportPath, JsonSerializer.Serialize(new {
            schema = "officeimo.pdf.invoice-comparison-evidence",
            schemaVersion = 2,
            generatedUtc = DateTimeOffset.UtcNow,
            environment = new {
                targetFramework = AppContext.TargetFrameworkName,
                runtime = RuntimeInformation.FrameworkDescription,
                os = RuntimeInformation.OSDescription,
                processArchitecture = RuntimeInformation.ProcessArchitecture.ToString()
            },
            source = new {
                commit = provenance.Commit,
                tree = provenance.Tree,
                worktreeClean = provenance.IsClean
            },
            contract = new {
                pages = 2,
                invoiceNumber = scenario.Invoice.Number,
                payableAmount = scenario.Calculation.PayableAmount,
                requiredText = scenario.RequiredText
            },
            artifacts
        }, new JsonSerializerOptions { WriteIndented = true }));
        Console.WriteLine(reportPath);
        return 0;
    }

    private static object Write(string output, string engine, byte[] pdf, InvoiceComparisonScenario scenario) {
        PdfInvoiceComparisonValidation.Validate(pdf, scenario, engine);
        string path = Path.Combine(output, engine + ".pdf");
        File.WriteAllBytes(path, pdf);

        PdfDocument document = PdfDocument.Load(pdf);
        IReadOnlyList<PdfPageRenderResult> rendered = document.Render.Pages("1-2", new PdfPageRenderOptions {
            Format = PdfPageRenderFormat.Png,
            Dpi = 120,
            MaxPages = 2,
            ContinueOnError = false
        });
        if (rendered.Count != 2) {
            throw new InvalidOperationException($"{engine} produced {rendered.Count} previews; expected 2.");
        }
        object[] previews = rendered.Select(page => WritePreview(output, engine, page)).ToArray();

        return new {
            engine,
            file = Path.GetFileName(path),
            bytes = pdf.Length,
            sha256 = Sha256(pdf),
            previews
        };
    }

    private static object WritePreview(string output, string engine, PdfPageRenderResult page) {
        byte[] bytes = page.Bytes
            ?? throw new InvalidOperationException($"{engine} page {page.PageNumber} did not produce preview bytes.");
        if (page.Width <= 0 || page.Height <= 0) {
            throw new InvalidOperationException($"{engine} page {page.PageNumber} produced invalid preview dimensions.");
        }
        string path = Path.Combine(output, $"{engine}-page-{page.PageNumber}.png");
        File.WriteAllBytes(path, bytes);
        return new {
            page = page.PageNumber,
            file = Path.GetFileName(path),
            width = page.Width,
            height = page.Height,
            bytes = bytes.Length,
            sha256 = Sha256(bytes),
            diagnostics = page.Diagnostics
        };
    }

    private static string? ReadOption(string[] args, string option) {
        for (int index = 1; index < args.Length - 1; index++) {
            if (string.Equals(args[index], option, StringComparison.OrdinalIgnoreCase)) return args[index + 1];
        }
        return null;
    }

    private static bool HasOption(string[] args, string option) =>
        args.Any(value => string.Equals(value, option, StringComparison.OrdinalIgnoreCase));

    private static string FindRepositoryRoot() {
        foreach (string seed in new[] { AppContext.BaseDirectory, Directory.GetCurrentDirectory() }) {
            string? current = Path.GetFullPath(seed);
            while (!string.IsNullOrWhiteSpace(current)) {
                if (File.Exists(Path.Combine(current, "OfficeIMO.sln"))) return current;
                current = Directory.GetParent(current)?.FullName;
            }
        }
        throw new DirectoryNotFoundException("Could not locate the OfficeIMO repository root.");
    }

    private static string Sha256(byte[] bytes) =>
        Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
}
