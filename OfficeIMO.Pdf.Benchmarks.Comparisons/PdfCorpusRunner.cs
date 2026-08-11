using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.RegularExpressions;
using UglyToad.PdfPig.DocumentLayoutAnalysis.TextExtractor;
using OfficePdfDocument = OfficeIMO.Pdf.PdfDocument;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

internal static partial class PdfCorpusRunner {
    private static readonly JsonSerializerOptions JsonOptions = new() {
        PropertyNameCaseInsensitive = true,
        WriteIndented = true
    };

    internal static async Task<int> RunAsync(string[] args) {
        string repositoryRoot = GetOption(args, "--repo-root") ?? FindRepositoryRoot();
        string manifestPath = Path.GetFullPath(
            GetOption(args, "--manifest") ??
            Path.Combine(repositoryRoot, "OfficeIMO.Pdf.Benchmarks.Comparisons", "Corpus", "pdf-corpus.json"));
        string outputDirectory = Path.GetFullPath(
            GetOption(args, "--output") ??
            Path.Combine(repositoryRoot, "Ignore", "Benchmarks", "PdfComparisons", "corpus"));
        bool download = HasOption(args, "--download");
        bool skipManipulation = HasOption(args, "--skip-manipulation");
        string? comPdfPath = GetOption(args, "--com-pdf");

        PdfCorpusManifest manifest = JsonSerializer.Deserialize<PdfCorpusManifest>(
            await File.ReadAllTextAsync(manifestPath).ConfigureAwait(false),
            JsonOptions) ?? throw new InvalidDataException("PDF corpus manifest is empty.");
        if (manifest.SchemaVersion != 1) {
            throw new InvalidDataException($"Unsupported PDF corpus schema {manifest.SchemaVersion}; expected 1.");
        }

        string filesDirectory = Path.Combine(outputDirectory, "files");
        string diagnosticsDirectory = Path.Combine(outputDirectory, "diagnostics");
        Directory.CreateDirectory(filesDirectory);
        Directory.CreateDirectory(diagnosticsDirectory);
        var entries = new List<PdfCorpusEntry>(manifest.Entries);
        if (!string.IsNullOrWhiteSpace(comPdfPath)) {
            entries.Add(new PdfCorpusEntry {
                Id = "microsoft-word-com-rich",
                SourceKind = "local",
                SourcePath = Path.GetFullPath(comPdfPath),
                Producer = "Microsoft Word COM export",
                License = "Generated local fixture",
                Tier = "large",
                ExpectedPages = 26,
                MinimumTokenRecall = 0.90D,
                Features = new List<string> { "tables", "chart", "native-word-smartart", "image", "links", "headers-footers", "office-com-export" },
                RequiredText = new List<string> {
                    "COM SMARTART INTEROPERABILITY PAGE",
                    "Collect",
                    "Validate",
                    "Publish"
                }
            });
        }
        string? only = GetOption(args, "--only");
        if (!string.IsNullOrWhiteSpace(only)) {
            var selectedIds = new HashSet<string>(
                only.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries),
                StringComparer.OrdinalIgnoreCase);
            entries = entries.Where(entry => selectedIds.Contains(entry.Id)).ToList();
            if (entries.Count == 0) {
                throw new ArgumentException($"No corpus entries matched --only {only}.", nameof(args));
            }
        }

        var results = new List<PdfCorpusResult>(entries.Count);
        using var client = new HttpClient { Timeout = TimeSpan.FromMinutes(5) };
        client.DefaultRequestHeaders.UserAgent.ParseAdd("OfficeIMO-Pdf-Corpus/1.0");
        foreach (PdfCorpusEntry entry in entries) {
            PdfCorpusResult result;
            try {
                string path = await ResolveArtifactAsync(entry, repositoryRoot, filesDirectory, download, client)
                    .ConfigureAwait(false);
                result = ValidateEntry(entry, path, skipManipulation, diagnosticsDirectory);
            } catch (Exception exception) {
                string unresolvedPath = entry.SourcePath ?? entry.Url ?? entry.Generator ?? string.Empty;
                result = new PdfCorpusResult(
                    entry.Id,
                    entry.Producer,
                    entry.Tier,
                    entry.SourceKind,
                    unresolvedPath,
                    0,
                    string.Empty,
                    entry.Features,
                    new PdfCorpusReadResult(false, "unavailable", 0, 0, 0, 0, exception.ToString()),
                    new PdfCorpusManipulationResult(false, "NotRun", 0, 0, 0, Array.Empty<string>(), "Read validation did not complete."));
            }

            results.Add(result);
            string outcome = !result.Read.Success || result.Manipulation.IsFailure
                ? "FAIL"
                : result.Manipulation.Status == "Blocked"
                    ? "BLOCK"
                    : "PASS";
            Console.WriteLine(
                $"{outcome,-5} " +
                $"{result.Id,-40} pages={result.Read.PageCount,4} " +
                $"recall={result.Read.TokenRecall:P1} bytes={result.Bytes:N0}");
            if (!result.Read.Success) {
                Console.WriteLine("  read: " + result.Read.Error);
            }
            if (!result.Manipulation.Success) {
                Console.WriteLine($"  manipulation [{result.Manipulation.Status}]: {result.Manipulation.Error}");
            }
        }

        var report = new PdfCorpusReport(
            1,
            DateTimeOffset.UtcNow,
            RuntimeInformation.FrameworkDescription,
            RuntimeInformation.OSDescription,
            results);
        string reportPath = Path.Combine(outputDirectory, "pdf-corpus-compatibility.json");
        await File.WriteAllTextAsync(reportPath, JsonSerializer.Serialize(report, JsonOptions)).ConfigureAwait(false);
        Console.WriteLine($"Corpus report: {reportPath}");
        return report.Success ? 0 : 2;
    }

    private static PdfCorpusResult ValidateEntry(
        PdfCorpusEntry entry,
        string path,
        bool skipManipulation,
        string diagnosticsDirectory) {
        byte[] bytes = File.ReadAllBytes(path);
        if (bytes.Length < 5 || !bytes.AsSpan(0, 4).SequenceEqual("%PDF"u8)) {
            throw new InvalidDataException($"{entry.Id} does not start with a PDF header.");
        }

        string sha256 = Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
        if (!string.IsNullOrWhiteSpace(entry.Sha256) &&
            !string.Equals(entry.Sha256, sha256, StringComparison.OrdinalIgnoreCase)) {
            throw new InvalidDataException(
                $"{entry.Id} SHA-256 mismatch. Expected {entry.Sha256}, observed {sha256}.");
        }

        IReadOnlyList<string> oraclePages;
        IReadOnlyList<string> officePages;
        string oracle;
        PdfCorpusReadResult read;
        try {
            (oraclePages, oracle) = ReadOracleByPage(bytes);
            OfficeIMO.Pdf.PdfReadOptions readOptions = new() {
                IncludeArtifactText = true
            };
            officePages = OfficePdfDocument.Open(bytes, readOptions).Read.TextByPage();
            File.WriteAllText(
                Path.Combine(diagnosticsDirectory, entry.Id + "." + oracle + ".txt"),
                string.Join("\n\f\n", oraclePages));
            File.WriteAllText(
                Path.Combine(diagnosticsDirectory, entry.Id + ".officeimo.txt"),
                string.Join("\n\f\n", officePages));
            if (entry.ExpectedPages.HasValue && oraclePages.Count != entry.ExpectedPages.Value) {
                throw new InvalidDataException(
                    $"{entry.Id} has {oraclePages.Count} pages; expected {entry.ExpectedPages.Value}.");
            }
            if (officePages.Count != oraclePages.Count) {
                throw new InvalidDataException(
                    $"OfficeIMO observed {officePages.Count} pages while PdfPig observed {oraclePages.Count}.");
            }
            ValidateRequiredText(entry, oraclePages, officePages);

            double recall = TokenRecall(oraclePages, officePages);
            if (recall < entry.MinimumTokenRecall) {
                WriteSpanDiagnostics(bytes, readOptions, diagnosticsDirectory, entry.Id);
                File.WriteAllText(
                    Path.Combine(diagnosticsDirectory, entry.Id + ".officeimo.debug.txt"),
                    OfficePdfDocument.Open(bytes).Debug(new OfficeIMO.Pdf.PdfDebuggerOptions {
                        IncludeDecodedStreamPreviews = true,
                        MaxDecodedStreamPreviewBytes = 64 * 1024
                    }).ToText());
                read = new PdfCorpusReadResult(
                    false,
                    oracle,
                    officePages.Count,
                    officePages.Sum(static text => text.Length),
                    oraclePages.Sum(static text => text.Length),
                    recall,
                    $"OfficeIMO token recall {recall:P2} is below the {entry.MinimumTokenRecall:P2} corpus threshold.");
                return CreateResult(entry, path, bytes.Length, sha256, read,
                    new PdfCorpusManipulationResult(false, "NotRun", 0, 0, 0, Array.Empty<string>(), "Read validation failed."));
            }

            read = new PdfCorpusReadResult(
                true,
                oracle,
                officePages.Count,
                officePages.Sum(static text => text.Length),
                oraclePages.Sum(static text => text.Length),
                recall,
                null);
        } catch (Exception exception) {
            read = new PdfCorpusReadResult(false, "unavailable", 0, 0, 0, 0, exception.ToString());
            return CreateResult(entry, path, bytes.Length, sha256, read,
                new PdfCorpusManipulationResult(false, "NotRun", 0, 0, 0, Array.Empty<string>(), "Read validation failed."));
        }

        PdfCorpusManipulationResult manipulation = skipManipulation
            ? new PdfCorpusManipulationResult(true, "Skipped", 0, 0, 0, Array.Empty<string>(), null)
            : ValidateManipulation(bytes, oraclePages);
        return CreateResult(entry, path, bytes.Length, sha256, read, manipulation);
    }

    private static void ValidateRequiredText(
        PdfCorpusEntry entry,
        IReadOnlyList<string> oraclePages,
        IReadOnlyList<string> officePages) {
        if (entry.RequiredText.Count == 0) {
            return;
        }

        string oracleText = string.Join('\n', oraclePages);
        string officeText = string.Join('\n', officePages);
        foreach (string requiredText in entry.RequiredText) {
            if (!oracleText.Contains(requiredText, StringComparison.OrdinalIgnoreCase)) {
                throw new InvalidDataException(
                    $"{entry.Id} independent extraction did not contain required text '{requiredText}'.");
            }
            if (!officeText.Contains(requiredText, StringComparison.OrdinalIgnoreCase)) {
                throw new InvalidDataException(
                    $"OfficeIMO did not recover required text '{requiredText}' from {entry.Id}.");
            }
        }
    }

    private static void WriteSpanDiagnostics(
        byte[] bytes,
        OfficeIMO.Pdf.PdfReadOptions? readOptions,
        string diagnosticsDirectory,
        string entryId) {
        const int maxSpans = 10_000;
        OfficeIMO.Pdf.PdfReadDocument document = OfficeIMO.Pdf.PdfReadDocument.Open(bytes, readOptions);
        using var writer = new StreamWriter(Path.Combine(diagnosticsDirectory, entryId + ".officeimo.spans.txt"));
        int written = 0;
        for (int pageIndex = 0; pageIndex < document.Pages.Count && written < maxSpans; pageIndex++) {
            writer.WriteLine($"PAGE {pageIndex + 1}");
            foreach (OfficeIMO.Pdf.PdfTextSpan span in document.Pages[pageIndex].GetTextSpans()) {
                writer.WriteLine(
                    $"X={span.X:F3} Y={span.Y:F3} A={span.Advance:F3} S={span.FontSize:F3} " +
                    $"F={span.FontResource} B={span.BaseFont ?? "-"} T={JsonSerializer.Serialize(span.Text)}");
                written++;
                if (written >= maxSpans) {
                    writer.WriteLine($"TRUNCATED after {maxSpans} spans");
                    break;
                }
            }
        }
    }

    private static PdfCorpusManipulationResult ValidateManipulation(byte[] source, IReadOnlyList<string> oraclePages) {
        try {
            int[] selectedPages = oraclePages.Count switch {
                0 => Array.Empty<int>(),
                1 => new[] { 1 },
                2 => new[] { 2, 1 },
                _ => new[] { oraclePages.Count, (oraclePages.Count + 1) / 2, 1 }
            };
            if (selectedPages.Length == 0) {
                throw new InvalidDataException("The source PDF contains no pages.");
            }

            byte[] selected = PdfManipulationEngines.SelectWithOfficeImo(source, selectedPages);
            IReadOnlyList<string> expected = selectedPages.Select(page => oraclePages[page - 1]).ToArray();
            ValidateManipulatedText("selection", selected, expected);

            byte[][] split = PdfManipulationEngines.SplitWithOfficeImo(selected, 1);
            if (split.Length != selectedPages.Length) {
                throw new InvalidDataException(
                    $"Split returned {split.Length} documents; expected {selectedPages.Length}.");
            }
            for (int index = 0; index < split.Length; index++) {
                ValidateManipulatedText("split page " + (index + 1), split[index], new[] { expected[index] });
            }

            byte[] merged = PdfManipulationEngines.MergeWithOfficeImo(split);
            ValidateManipulatedText("merge", merged, expected);
            return new PdfCorpusManipulationResult(
                true,
                "Passed",
                selectedPages.Length,
                split.Length,
                expected.Count,
                Array.Empty<string>(),
                null);
        } catch (OfficeIMO.Pdf.PdfMutationBlockedException exception) {
            return new PdfCorpusManipulationResult(
                false,
                "Blocked",
                0,
                0,
                0,
                exception.Plan.BlockerCodes,
                exception.Message);
        } catch (Exception exception) {
            return new PdfCorpusManipulationResult(
                false,
                "Failed",
                0,
                0,
                0,
                Array.Empty<string>(),
                exception.ToString());
        }
    }

    private static void ValidateManipulatedText(string operation, byte[] pdf, IReadOnlyList<string> expectedPages) {
        (IReadOnlyList<string> actualPages, _) = ReadOracleByPage(pdf);
        if (actualPages.Count != expectedPages.Count) {
            throw new InvalidDataException(
                $"OfficeIMO {operation} produced {actualPages.Count} pages; expected {expectedPages.Count}.");
        }

        double recall = TokenRecall(expectedPages, actualPages);
        if (recall < 0.98D) {
            throw new InvalidDataException(
                $"OfficeIMO {operation} retained only {recall:P2} of the independently extracted source tokens.");
        }
    }

    private static PdfCorpusResult CreateResult(
        PdfCorpusEntry entry,
        string path,
        long bytes,
        string sha256,
        PdfCorpusReadResult read,
        PdfCorpusManipulationResult manipulation) =>
        new(entry.Id, entry.Producer, entry.Tier, entry.SourceKind, path, bytes, sha256, entry.Features, read, manipulation);

    private static (IReadOnlyList<string> Pages, string Oracle) ReadOracleByPage(byte[] bytes) {
        try {
            using var stream = new MemoryStream(bytes, writable: false);
            using UglyToad.PdfPig.PdfDocument document = UglyToad.PdfPig.PdfDocument.Open(stream);
            return (
                document.GetPages().Select(static page => ContentOrderTextExtractor.GetText(page)).ToArray(),
                "pdfpig");
        } catch {
            return (ReadITextByPage(bytes), "itext");
        }
    }

    private static IReadOnlyList<string> ReadITextByPage(byte[] bytes) {
        using var stream = new MemoryStream(bytes, writable: false);
        using var reader = new iText.Kernel.Pdf.PdfReader(stream);
        using var document = new iText.Kernel.Pdf.PdfDocument(reader);
        var pages = new string[document.GetNumberOfPages()];
        for (int page = 1; page <= pages.Length; page++) {
            pages[page - 1] = iText.Kernel.Pdf.Canvas.Parser.PdfTextExtractor.GetTextFromPage(
                document.GetPage(page),
                new iText.Kernel.Pdf.Canvas.Parser.Listener.LocationTextExtractionStrategy());
        }
        return pages;
    }

    internal static double TokenRecall(IReadOnlyList<string> expectedPages, IReadOnlyList<string> actualPages) {
        long expectedTokenCount = 0;
        long matchedTokenCount = 0;
        int pages = Math.Min(expectedPages.Count, actualPages.Count);
        for (int index = 0; index < pages; index++) {
            string[] expected = Tokens(expectedPages[index]);
            string[] actual = Tokens(actualPages[index]);
            expectedTokenCount += expected.Length;
            var actualCounts = new Dictionary<string, int>(StringComparer.Ordinal);
            foreach (string token in actual) {
                actualCounts.TryGetValue(token, out int count);
                actualCounts[token] = count + 1;
            }
            foreach (string token in expected) {
                if (actualCounts.TryGetValue(token, out int count) && count > 0) {
                    matchedTokenCount++;
                    actualCounts[token] = count - 1;
                }
            }
        }

        return expectedTokenCount == 0 ? 1D : matchedTokenCount / (double)expectedTokenCount;
    }

    private static string[] Tokens(string value) => TokenRegex()
        .Matches(value.Normalize())
        .Select(static match => match.Value.ToUpperInvariant())
        .ToArray();

    private static async Task<string> ResolveArtifactAsync(
        PdfCorpusEntry entry,
        string repositoryRoot,
        string filesDirectory,
        bool download,
        HttpClient client) {
        switch (entry.SourceKind.ToLowerInvariant()) {
            case "repository":
                return Path.GetFullPath(Path.Combine(repositoryRoot,
                    entry.SourcePath ?? throw new InvalidDataException($"{entry.Id} has no sourcePath.")));
            case "local":
                return Path.GetFullPath(entry.SourcePath ?? throw new InvalidDataException($"{entry.Id} has no sourcePath."));
            case "generated":
                return GenerateArtifact(entry, repositoryRoot, filesDirectory);
            case "download":
                return await ResolveDownloadedArtifactAsync(entry, filesDirectory, download, client)
                    .ConfigureAwait(false);
            default:
                throw new InvalidDataException($"Unknown PDF corpus source kind '{entry.SourceKind}'.");
        }
    }

    private static async Task<string> ResolveDownloadedArtifactAsync(
        PdfCorpusEntry entry,
        string filesDirectory,
        bool download,
        HttpClient client) {
        string target = Path.Combine(filesDirectory, entry.Id + ".pdf");
        if (File.Exists(target)) {
            byte[] cachedPayload = await File.ReadAllBytesAsync(target).ConfigureAwait(false);
            string? cacheError = GetDownloadValidationError(entry, cachedPayload);
            if (cacheError is null) {
                return target;
            }

            if (!download) {
                throw new InvalidDataException(
                    $"Cached corpus artifact {entry.Id} is invalid: {cacheError} Re-run with --download to replace it.");
            }
        } else if (!download) {
            throw new FileNotFoundException(
                $"{entry.Id} is not prepared. Re-run with --download.", target);
        }

        byte[] payload = await client.GetByteArrayAsync(
            entry.Url ?? throw new InvalidDataException($"{entry.Id} has no URL."))
            .ConfigureAwait(false);
        string? validationError = GetDownloadValidationError(entry, payload);
        if (validationError is not null) {
            throw new InvalidDataException(
                $"Downloaded corpus artifact {entry.Id} is invalid: {validationError}");
        }

        string temporaryPath = target + "." + Guid.NewGuid().ToString("N") + ".download";
        try {
            await File.WriteAllBytesAsync(temporaryPath, payload).ConfigureAwait(false);
            File.Move(temporaryPath, target, overwrite: true);
        } finally {
            if (File.Exists(temporaryPath)) {
                File.Delete(temporaryPath);
            }
        }

        return target;
    }

    private static string? GetDownloadValidationError(PdfCorpusEntry entry, byte[] payload) {
        if (payload.Length < 5 || !payload.AsSpan(0, 4).SequenceEqual("%PDF"u8)) {
            return "the payload does not start with a PDF header.";
        }

        if (string.IsNullOrWhiteSpace(entry.Sha256)) {
            return "the download manifest does not declare a SHA-256 digest.";
        }

        string observedSha256 = Convert.ToHexString(SHA256.HashData(payload)).ToLowerInvariant();
        return string.Equals(entry.Sha256, observedSha256, StringComparison.OrdinalIgnoreCase)
            ? null
            : $"SHA-256 mismatch. Expected {entry.Sha256}, observed {observedSha256}.";
    }

    private static string GenerateArtifact(PdfCorpusEntry entry, string repositoryRoot, string filesDirectory) =>
        entry.Generator switch {
            "officeimo-word-rich" => RichWordPdfCorpusGenerator.Generate(repositoryRoot, filesDirectory).PdfPath,
            "officeimo-native-large" => GenerateLargeNativePdf(filesDirectory),
            _ => throw new InvalidDataException($"Unknown PDF corpus generator '{entry.Generator}'.")
        };

    private static string GenerateLargeNativePdf(string filesDirectory) {
        string path = Path.Combine(filesDirectory, "officeimo-native-large-500-page.pdf");
        var scenario = new PdfBenchmarkScenario(
            PdfBenchmarkScale.High,
            "Large PDF reader and manipulation corpus",
            PageCount: 500,
            RowsPerPage: 4,
            ParagraphsPerPage: 1);
        File.WriteAllBytes(path, OfficeImoPdfGenerator.Generate(scenario));
        return path;
    }

    private static bool HasOption(string[] args, string option) =>
        args.Any(value => string.Equals(value, option, StringComparison.OrdinalIgnoreCase));

    private static string? GetOption(string[] args, string option) {
        for (int index = 1; index < args.Length - 1; index++) {
            if (string.Equals(args[index], option, StringComparison.OrdinalIgnoreCase)) {
                return args[index + 1];
            }
        }
        return null;
    }

    private static string FindRepositoryRoot() {
        DirectoryInfo? directory = new(Environment.CurrentDirectory);
        while (directory != null) {
            if (File.Exists(Path.Combine(directory.FullName, "OfficeIMO.sln"))) {
                return directory.FullName;
            }
            directory = directory.Parent;
        }
        throw new DirectoryNotFoundException("Could not locate the OfficeIMO repository root.");
    }

    [GeneratedRegex(@"[\p{L}\p{Nd}]{3,}", RegexOptions.CultureInvariant)]
    private static partial Regex TokenRegex();
}
