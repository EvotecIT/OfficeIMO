using System.Text.Json;
using OfficeIMO.Markdown;

namespace OfficeIMO.Tests.MarkdownSuite;

internal static class CommonMarkInventory {
    public static CommonMarkInventoryReport Build(string officialSpecPath, string pinnedFixturePath) {
        var official = LoadExamples(officialSpecPath);
        var pinned = LoadExamples(pinnedFixturePath);
        var pinnedIds = new HashSet<int>(pinned.Select(static example => example.Example));
        var sectionOrder = official
            .Select(static example => example.Section)
            .Distinct(StringComparer.Ordinal)
            .ToArray();

        var entries = new List<CommonMarkInventoryEntry>(official.Count);
        foreach (var example in official) {
            bool isPinned = pinnedIds.Contains(example.Example);
            entries.Add(Evaluate(example, isPinned));
        }

        return new CommonMarkInventoryReport(sectionOrder, entries);
    }

    private static CommonMarkInventoryEntry Evaluate(CommonMarkSpecExample example, bool isPinned) {
        try {
            var result = OfficeIMO.Markdown.MarkdownReader.ParseWithSyntaxTree(example.Markdown, MarkdownReaderOptions.CreateCommonMarkProfile());
            string actual = result.Document.ToHtmlFragment(CommonMarkHtmlComparison.CreatePlainHtmlOptions());
            bool matches = CommonMarkHtmlComparison.Normalize(example.Html) == CommonMarkHtmlComparison.Normalize(actual);

            return new CommonMarkInventoryEntry(
                example.Example,
                example.Section,
                isPinned,
                matches ? CommonMarkInventoryStatus.Passing : CommonMarkInventoryStatus.Failing,
                ClassifyCluster(example.Section),
                matches ? null : DescribeMismatch(example.Html, actual));
        } catch (Exception ex) {
            return new CommonMarkInventoryEntry(
                example.Example,
                example.Section,
                isPinned,
                CommonMarkInventoryStatus.Failing,
                ClassifyCluster(example.Section),
                ex.GetType().Name + ": " + ex.Message);
        }
    }

    private static string DescribeMismatch(string expectedHtml, string actualHtml) {
        string expected = CommonMarkHtmlComparison.Normalize(expectedHtml);
        string actual = CommonMarkHtmlComparison.Normalize(actualHtml);
        return "expected " + Abbreviate(expected) + "; actual " + Abbreviate(actual);
    }

    private static string Abbreviate(string value) {
        value = value.Replace("\r", "\\r").Replace("\n", "\\n");
        return value.Length <= 120 ? value : value.Substring(0, 117) + "...";
    }

    private static string ClassifyCluster(string section) {
        return section switch {
            "Entity and numeric character references" => "CommonMark entity decoder",
            "HTML blocks" or "Raw HTML" => "HTML block/raw HTML grammar",
            "Code spans" => "Code span normalization and precedence",
            "Links" or "Link reference definitions" or "Images" => "Link/image/reference grammar",
            "Emphasis and strong emphasis" => "Emphasis delimiter algorithm",
            "Tabs" or "Indented code blocks" or "Block quotes" or "List items" => "Container indentation and continuation",
            "Autolinks" => "Autolink grammar",
            "Hard line breaks" or "Soft line breaks" or "Backslash escapes" or "Inlines" or "Precedence" => "Inline precedence and line-break grammar",
            _ => "Baseline block/text coverage"
        };
    }

    private static IReadOnlyList<CommonMarkSpecExample> LoadExamples(string path) {
        string json = File.ReadAllText(path);
        var examples = JsonSerializer.Deserialize<List<CommonMarkSpecExample>>(json, new JsonSerializerOptions {
            PropertyNameCaseInsensitive = true
        });

        if (examples == null || examples.Count == 0) {
            throw new InvalidOperationException("CommonMark examples were not loaded from " + path + ".");
        }

        return examples;
    }
}

internal sealed class CommonMarkInventoryReport {
    public CommonMarkInventoryReport(IReadOnlyList<string> sectionOrder, IReadOnlyList<CommonMarkInventoryEntry> entries) {
        SectionOrder = sectionOrder;
        Entries = entries;
    }

    public IReadOnlyList<string> SectionOrder { get; }
    public IReadOnlyList<CommonMarkInventoryEntry> Entries { get; }

    public int Total => Entries.Count;
    public int Pinned => Entries.Count(static entry => entry.IsPinned);
    public int PassingPinned => Entries.Count(static entry => entry.IsPinned && entry.Status == CommonMarkInventoryStatus.Passing);
    public int PassingUnpinned => Entries.Count(static entry => !entry.IsPinned && entry.Status == CommonMarkInventoryStatus.Passing);
    public int Failing => Entries.Count(static entry => entry.Status == CommonMarkInventoryStatus.Failing);
    public int IntentionalDeviations => Entries.Count(static entry => entry.Status == CommonMarkInventoryStatus.IntentionalDeviation);

    public IEnumerable<CommonMarkSectionSummary> EnumerateSectionSummaries() {
        foreach (string section in SectionOrder) {
            var entries = Entries.Where(entry => string.Equals(entry.Section, section, StringComparison.Ordinal)).ToArray();
            yield return new CommonMarkSectionSummary(
                section,
                entries.Length,
                entries.Count(static entry => entry.IsPinned),
                entries.Count(static entry => entry.IsPinned && entry.Status == CommonMarkInventoryStatus.Passing),
                entries.Count(static entry => !entry.IsPinned && entry.Status == CommonMarkInventoryStatus.Passing),
                entries.Count(static entry => entry.Status == CommonMarkInventoryStatus.Failing),
                entries.Count(static entry => entry.Status == CommonMarkInventoryStatus.IntentionalDeviation));
        }
    }

    public IEnumerable<CommonMarkFailureClusterSummary> EnumerateFailureClusters() {
        return Entries
            .Where(static entry => entry.Status == CommonMarkInventoryStatus.Failing)
            .GroupBy(static entry => entry.Cluster, StringComparer.Ordinal)
            .Select(static group => new CommonMarkFailureClusterSummary(
                group.Key,
                group.Count(),
                string.Join(", ", group.Select(static entry => entry.Section).Distinct(StringComparer.Ordinal).OrderBy(static section => section, StringComparer.Ordinal)),
                group.Select(static entry => entry.Example).OrderBy(static example => example).Take(12).ToArray()))
            .OrderByDescending(static cluster => cluster.Count)
            .ThenBy(static cluster => cluster.Cluster, StringComparer.Ordinal);
    }
}

internal sealed class CommonMarkInventoryEntry {
    public CommonMarkInventoryEntry(int example, string section, bool isPinned, CommonMarkInventoryStatus status, string cluster, string? detail) {
        Example = example;
        Section = section;
        IsPinned = isPinned;
        Status = status;
        Cluster = cluster;
        Detail = detail;
    }

    public int Example { get; }
    public string Section { get; }
    public bool IsPinned { get; }
    public CommonMarkInventoryStatus Status { get; }
    public string Cluster { get; }
    public string? Detail { get; }
}

internal sealed class CommonMarkSectionSummary {
    public CommonMarkSectionSummary(string section, int total, int pinned, int passingPinned, int passingUnpinned, int failing, int intentionalDeviations) {
        Section = section;
        Total = total;
        Pinned = pinned;
        PassingPinned = passingPinned;
        PassingUnpinned = passingUnpinned;
        Failing = failing;
        IntentionalDeviations = intentionalDeviations;
    }

    public string Section { get; }
    public int Total { get; }
    public int Pinned { get; }
    public int PassingPinned { get; }
    public int PassingUnpinned { get; }
    public int Failing { get; }
    public int IntentionalDeviations { get; }
}

internal sealed class CommonMarkFailureClusterSummary {
    public CommonMarkFailureClusterSummary(string cluster, int count, string sections, IReadOnlyList<int> examples) {
        Cluster = cluster;
        Count = count;
        Sections = sections;
        Examples = examples;
    }

    public string Cluster { get; }
    public int Count { get; }
    public string Sections { get; }
    public IReadOnlyList<int> Examples { get; }
}

internal enum CommonMarkInventoryStatus {
    Passing,
    Failing,
    IntentionalDeviation
}

internal sealed class CommonMarkSpecExample {
    public int Example { get; set; }
    public string Section { get; set; } = string.Empty;
    public string Markdown { get; set; } = string.Empty;
    public string Html { get; set; } = string.Empty;
}
