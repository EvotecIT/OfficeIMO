using System.Text.Json;
using OfficeIMO.Pdf;
using Xunit;
using OfficePdfDocument = OfficeIMO.Pdf.PdfDocument;

namespace OfficeIMO.Tests.Pdf;

/// <summary>
/// Opens, extracts and renders the committed files of the independent-producer corpus.
/// Labels come from <c>pdf-corpus-extended.json</c>; larger files there stay download-only.
/// </summary>
public sealed class PdfIndependentCorpusTests {
    private const string ManifestPath = "OfficeIMO.Pdf.Benchmarks.Comparisons/Corpus/pdf-corpus-extended.json";
    private const string CorpusFolder = "OfficeIMO.TestAssets/Pdf/IndependentCorpus";
    // These budgets catch hangs, not slowness; they stay well clear of shared-runner load.
    private static readonly TimeSpan FileBudget = TimeSpan.FromSeconds(60);
    private static readonly TimeSpan RenderBudget = TimeSpan.FromSeconds(45);
    // Large scanned pages are rendered downscaled rather than skipped.
    private const int MaximumRenderDimension = 1600;

    // Current defects from Docs/benchmarks/pdf-independent-corpus-2026-11-10.md. Each entry must still fail at
    // every listed stage and nowhere else; the test fails once a defect is fixed so the entry gets removed.
    private static readonly Dictionary<string, KnownFailure> KnownFailures = new(StringComparer.Ordinal) {
        ["qpdf-inline-images"] = new(
            "Ghostscript 5.50 inline images at 600 dpi exceed the managed raster limits.",
            Stage.Render),
        ["nara-jfk-104-10003-10041"] = new(
            "The 16.4-megapixel scan on a letter-size page exceeds the 16-megapixel display limit, " +
            "although the requested output is at most 1600 px.",
            Stage.Render)
    };

    private static readonly Lazy<IReadOnlyDictionary<string, CorpusEntry>> Entries = new(LoadEntries);

    public static TheoryData<string> CommittedFiles() {
        var data = new TheoryData<string>();
        foreach (string id in Entries.Value.Keys.OrderBy(id => id, StringComparer.Ordinal)) data.Add(id);
        return data;
    }

    [Fact]
    public void CommittedFiles_MatchTheManifest() {
        string folder = Path.Combine(RepositoryTestPaths.Find(), CorpusFolder);
        foreach (CorpusEntry entry in Entries.Value.Values) {
            Assert.True(File.Exists(entry.FilePath), "Missing committed corpus file: " + entry.Id);
            Assert.Equal(entry.Sha256, PdfArtifactFingerprint.ComputeSha256(File.ReadAllBytes(entry.FilePath)));
        }

        string[] committed = Directory.GetFiles(folder, "*.pdf")
            .Select(path => Path.GetFileNameWithoutExtension(path)!)
            .OrderBy(id => id, StringComparer.Ordinal)
            .ToArray();
        Assert.Equal(Entries.Value.Keys.OrderBy(id => id, StringComparer.Ordinal), committed);
        Assert.All(KnownFailures.Keys, id => Assert.Contains(id, Entries.Value.Keys));
    }

    [Theory]
    [MemberData(nameof(CommittedFiles))]
    public async Task CommittedFile_LoadsExtractsTextAndRendersFirstPage(string id) {
        CorpusEntry entry = Entries.Value[id];
        using var cancellation = new CancellationTokenSource(FileBudget);
        Task<IReadOnlyList<StageFailure>> work = Task.Run(() => Exercise(entry, cancellation.Token));
        Task finished = await Task.WhenAny(work, Task.Delay(FileBudget + TimeSpan.FromSeconds(5)));
        Assert.True(finished == work, id + " did not finish within " + FileBudget.TotalSeconds + " s.");
        IReadOnlyList<StageFailure> failures = await work;

        if (KnownFailures.TryGetValue(id, out KnownFailure? known)) {
            foreach (Stage stage in known.Stages) {
                Assert.True(
                    failures.Any(failure => failure.Stage == stage),
                    id + " no longer fails at " + stage + " (" + known.Defect + "). Update its KnownFailures entry.");
            }
            StageFailure[] unexpected = failures.Where(failure => !known.Stages.Contains(failure.Stage)).ToArray();
            Assert.True(unexpected.Length == 0, id + " failed beyond its known defect: " + Describe(unexpected));
        } else {
            Assert.True(failures.Count == 0, id + " failed: " + Describe(failures));
        }
    }

    private static IReadOnlyList<StageFailure> Exercise(CorpusEntry entry, CancellationToken cancellationToken) {
        var failures = new List<StageFailure>();
        using var input = new MemoryStream(File.ReadAllBytes(entry.FilePath), writable: false);
        OfficePdfDocument document;
        int pageCount;
        try {
            document = OfficePdfDocument.Load(input, new PdfLoadOptions { Password = entry.Password }, cancellationToken);
            pageCount = document.Pages.Count;
        } catch (Exception exception) {
            failures.Add(new StageFailure(Stage.Load, Summarize(exception)));
            return failures;
        }

        using (document) {
            if (entry.ExpectedPages is int expectedPages && pageCount != expectedPages) {
                failures.Add(new StageFailure(Stage.PageCount, pageCount + " pages; expected " + expectedPages));
            }

            try {
                string text = document.ExtractText(cancellationToken);
                foreach (string expected in entry.ExpectedText) {
                    if (text.IndexOf(expected, StringComparison.Ordinal) < 0) {
                        failures.Add(new StageFailure(Stage.ExpectedText, "missing '" + expected + "'"));
                    }
                }
            } catch (Exception exception) {
                failures.Add(new StageFailure(Stage.Text, Summarize(exception)));
            }

            if (pageCount > 0) {
                try {
                    PdfPageRenderResult page = document.Render.DisplayPage(
                        1, new PdfPageDisplayOptions { MaximumDimension = MaximumRenderDimension, Timeout = RenderBudget }, cancellationToken);
                    if (!page.Succeeded && page.Diagnostics.Count == 0 && page.CapabilityDiagnostics.Count == 0) {
                        failures.Add(new StageFailure(Stage.Render, "no image and no diagnostics"));
                    }
                } catch (Exception exception) {
                    failures.Add(new StageFailure(Stage.Render, Summarize(exception)));
                }
            }
        }

        return failures;
    }

    private static IReadOnlyDictionary<string, CorpusEntry> LoadEntries() {
        string root = RepositoryTestPaths.Find();
        using JsonDocument manifest = JsonDocument.Parse(File.ReadAllBytes(Path.Combine(root, ManifestPath)));
        var entries = new Dictionary<string, CorpusEntry>(StringComparer.Ordinal);
        foreach (JsonElement item in manifest.RootElement.GetProperty("entries").EnumerateArray()) {
            if (!item.TryGetProperty("sourcePath", out JsonElement sourcePath)) continue;
            string id = item.GetProperty("id").GetString()!;
            entries.Add(id, new CorpusEntry(
                id,
                Path.Combine(root, sourcePath.GetString()!),
                item.GetProperty("sha256").GetString()!,
                item.TryGetProperty("password", out JsonElement password) ? password.GetString() : null,
                item.TryGetProperty("expectedPages", out JsonElement pages) ? pages.GetInt32() : null,
                item.TryGetProperty("expectedText", out JsonElement text)
                    ? text.EnumerateArray().Select(value => value.GetString()!).ToArray()
                    : Array.Empty<string>()));
        }
        return entries;
    }

    private static string Summarize(Exception exception) {
        string message = exception.Message;
        int lineBreak = message.IndexOfAny(new[] { '\r', '\n' });
        return exception.GetType().Name + ": " + (lineBreak < 0 ? message : message.Substring(0, lineBreak));
    }

    private static string Describe(IEnumerable<StageFailure> failures) =>
        string.Join("; ", failures.Select(failure => failure.Stage + " - " + failure.Detail));

    private enum Stage {
        Load,
        PageCount,
        Text,
        ExpectedText,
        Render
    }

    private sealed record CorpusEntry(
        string Id,
        string FilePath,
        string Sha256,
        string? Password,
        int? ExpectedPages,
        IReadOnlyList<string> ExpectedText);

    private sealed record StageFailure(Stage Stage, string Detail);

    private sealed class KnownFailure {
        internal KnownFailure(string defect, params Stage[] stages) {
            Defect = defect;
            Stages = stages;
        }

        internal string Defect { get; }
        internal Stage[] Stages { get; }
    }
}
