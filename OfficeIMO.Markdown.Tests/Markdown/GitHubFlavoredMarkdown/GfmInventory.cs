using System.Text.Json;
using OfficeIMO.Markdown;

namespace OfficeIMO.Tests.MarkdownSuite;

internal static class GfmInventory {
    public static GfmInventoryReport Build(string fixturePath) {
        var fixtures = LoadFixtures(fixturePath);
        var sectionOrder = fixtures
            .Select(static fixture => fixture.Section)
            .Distinct(StringComparer.Ordinal)
            .ToArray();

        var entries = new List<GfmInventoryEntry>(fixtures.Count);
        for (int i = 0; i < fixtures.Count; i++) {
            entries.Add(Evaluate(i + 1, fixtures[i]));
        }

        return new GfmInventoryReport(sectionOrder, entries);
    }

    private static GfmInventoryEntry Evaluate(int index, GfmExampleFixture fixture) {
        try {
            var result = OfficeIMO.Markdown.MarkdownReader.ParseWithSyntaxTree(fixture.Markdown, MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile());
            string actual = result.Document.ToHtmlFragment(GfmHtmlComparison.CreatePlainHtmlOptions());
            bool matches = GfmHtmlComparison.Normalize(fixture.Html) == GfmHtmlComparison.Normalize(actual);

            return new GfmInventoryEntry(
                index,
                fixture.Source,
                fixture.Section,
                IsUpstreamSource(fixture.Source),
                matches ? GfmInventoryStatus.Passing : GfmInventoryStatus.Failing,
                ClassifyCluster(fixture.Section),
                matches ? null : DescribeMismatch(fixture.Html, actual));
        } catch (Exception ex) {
            return new GfmInventoryEntry(
                index,
                fixture.Source,
                fixture.Section,
                IsUpstreamSource(fixture.Source),
                GfmInventoryStatus.Failing,
                ClassifyCluster(fixture.Section),
                ex.GetType().Name + ": " + ex.Message);
        }
    }

    private static bool IsUpstreamSource(string source) =>
        source.StartsWith("github/cmark-gfm ", StringComparison.Ordinal);

    private static string DescribeMismatch(string expectedHtml, string actualHtml) {
        string expected = GfmHtmlComparison.Normalize(expectedHtml);
        string actual = GfmHtmlComparison.Normalize(actualHtml);
        return "expected " + Abbreviate(expected) + "; actual " + Abbreviate(actual);
    }

    private static string Abbreviate(string value) {
        value = value.Replace("\r", "\\r").Replace("\n", "\\n");
        return value.Length <= 120 ? value : value.Substring(0, 117) + "...";
    }

    private static string ClassifyCluster(string section) {
        return section switch {
            "Tables" => "GFM table grammar",
            "Task lists" => "GFM task-list grammar",
            "Autolinks" => "GFM autolink grammar",
            "Strikethroughs" => "GFM strikethrough delimiter algorithm",
            "HTML tag filter" => "GFM HTML tag filter",
            "Footnotes" => "GFM footnote rendering",
            "Interop" => "GFM extension interaction",
            _ => "GFM extension coverage"
        };
    }

    private static IReadOnlyList<GfmExampleFixture> LoadFixtures(string path) {
        string json = File.ReadAllText(path);
        var fixtures = JsonSerializer.Deserialize<List<GfmExampleFixture>>(json, new JsonSerializerOptions {
            PropertyNameCaseInsensitive = true
        });

        if (fixtures == null || fixtures.Count == 0) {
            throw new InvalidOperationException("GFM fixtures were not loaded from " + path + ".");
        }

        return fixtures;
    }
}

internal sealed class GfmInventoryReport {
    public GfmInventoryReport(IReadOnlyList<string> sectionOrder, IReadOnlyList<GfmInventoryEntry> entries) {
        SectionOrder = sectionOrder;
        Entries = entries;
    }

    public IReadOnlyList<string> SectionOrder { get; }
    public IReadOnlyList<GfmInventoryEntry> Entries { get; }

    public int Total => Entries.Count;
    public int UpstreamTracked => Entries.Count(static entry => entry.IsUpstream);
    public int Supplements => Entries.Count(static entry => !entry.IsUpstream);
    public int Passing => Entries.Count(static entry => entry.Status == GfmInventoryStatus.Passing);
    public int Failing => Entries.Count(static entry => entry.Status == GfmInventoryStatus.Failing);
    public int IntentionalDeviations => Entries.Count(static entry => entry.Status == GfmInventoryStatus.IntentionalDeviation);

    public IEnumerable<GfmSectionSummary> EnumerateSectionSummaries() {
        foreach (string section in SectionOrder) {
            var entries = Entries.Where(entry => string.Equals(entry.Section, section, StringComparison.Ordinal)).ToArray();
            yield return new GfmSectionSummary(
                section,
                entries.Length,
                entries.Count(static entry => entry.IsUpstream),
                entries.Count(static entry => !entry.IsUpstream),
                entries.Count(static entry => entry.Status == GfmInventoryStatus.Passing),
                entries.Count(static entry => entry.Status == GfmInventoryStatus.Failing),
                entries.Count(static entry => entry.Status == GfmInventoryStatus.IntentionalDeviation));
        }
    }

    public IEnumerable<GfmSourceSummary> EnumerateSourceSummaries() {
        return Entries
            .GroupBy(static entry => entry.Source, StringComparer.Ordinal)
            .Select(static group => new GfmSourceSummary(
                group.Key,
                group.Count(),
                group.Count(static entry => entry.Status == GfmInventoryStatus.Passing),
                group.Count(static entry => entry.Status == GfmInventoryStatus.Failing)))
            .OrderBy(static source => source.Source, StringComparer.Ordinal);
    }

    public IEnumerable<GfmFailureClusterSummary> EnumerateFailureClusters() {
        return Entries
            .Where(static entry => entry.Status == GfmInventoryStatus.Failing)
            .GroupBy(static entry => entry.Cluster, StringComparer.Ordinal)
            .Select(static group => new GfmFailureClusterSummary(
                group.Key,
                group.Count(),
                string.Join(", ", group.Select(static entry => entry.Section).Distinct(StringComparer.Ordinal).OrderBy(static section => section, StringComparer.Ordinal)),
                group.Select(static entry => entry.Index).OrderBy(static index => index).Take(12).ToArray()))
            .OrderByDescending(static cluster => cluster.Count)
            .ThenBy(static cluster => cluster.Cluster, StringComparer.Ordinal);
    }
}

internal sealed class GfmInventoryEntry {
    public GfmInventoryEntry(int index, string source, string section, bool isUpstream, GfmInventoryStatus status, string cluster, string? detail) {
        Index = index;
        Source = source;
        Section = section;
        IsUpstream = isUpstream;
        Status = status;
        Cluster = cluster;
        Detail = detail;
    }

    public int Index { get; }
    public string Source { get; }
    public string Section { get; }
    public bool IsUpstream { get; }
    public GfmInventoryStatus Status { get; }
    public string Cluster { get; }
    public string? Detail { get; }
}

internal sealed class GfmSectionSummary {
    public GfmSectionSummary(string section, int total, int upstream, int supplements, int passing, int failing, int intentionalDeviations) {
        Section = section;
        Total = total;
        Upstream = upstream;
        Supplements = supplements;
        Passing = passing;
        Failing = failing;
        IntentionalDeviations = intentionalDeviations;
    }

    public string Section { get; }
    public int Total { get; }
    public int Upstream { get; }
    public int Supplements { get; }
    public int Passing { get; }
    public int Failing { get; }
    public int IntentionalDeviations { get; }
}

internal sealed class GfmSourceSummary {
    public GfmSourceSummary(string source, int total, int passing, int failing) {
        Source = source;
        Total = total;
        Passing = passing;
        Failing = failing;
    }

    public string Source { get; }
    public int Total { get; }
    public int Passing { get; }
    public int Failing { get; }
}

internal sealed class GfmFailureClusterSummary {
    public GfmFailureClusterSummary(string cluster, int count, string sections, IReadOnlyList<int> indexes) {
        Cluster = cluster;
        Count = count;
        Sections = sections;
        Indexes = indexes;
    }

    public string Cluster { get; }
    public int Count { get; }
    public string Sections { get; }
    public IReadOnlyList<int> Indexes { get; }
}

internal enum GfmInventoryStatus {
    Passing,
    Failing,
    IntentionalDeviation
}

public sealed class GfmExampleFixture {
    public string Source { get; set; } = string.Empty;
    public string Section { get; set; } = string.Empty;
    public string Markdown { get; set; } = string.Empty;
    public string Html { get; set; } = string.Empty;
    public string[] TopLevelKinds { get; set; } = [];
    public MarkdownSpecSyntaxAssertionFixture[] SyntaxAssertions { get; set; } = [];
}
