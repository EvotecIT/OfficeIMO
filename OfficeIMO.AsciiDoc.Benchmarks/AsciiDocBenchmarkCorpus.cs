using System.Collections.Concurrent;
using System.Text;

namespace OfficeIMO.AsciiDoc.Benchmarks;

internal sealed record AsciiDocBenchmarkFixture(string Scale, string Source, int SectionCount, int RecordCount);

internal static class AsciiDocBenchmarkCorpus {
    private static readonly ConcurrentDictionary<string, AsciiDocBenchmarkFixture> Fixtures = new(StringComparer.OrdinalIgnoreCase);
    internal static IReadOnlyList<string> Scales { get; } = ["Small", "Normal", "Large"];

    internal static AsciiDocBenchmarkFixture Get(string scale) {
        string selected = Scales.FirstOrDefault(value => string.Equals(value, scale, StringComparison.OrdinalIgnoreCase))
            ?? throw new ArgumentException("Unknown AsciiDoc benchmark scale: " + scale, nameof(scale));
        return Fixtures.GetOrAdd(selected, Create);
    }

    private static AsciiDocBenchmarkFixture Create(string scale) {
        int sectionCount = scale switch { "Small" => 2, "Normal" => 100, _ => 1_000 };
        const int recordsPerSection = 4;
        var source = new StringBuilder(sectionCount * 750);
        source.AppendLine("= OfficeIMO deterministic AsciiDoc workload");
        source.AppendLine(":toc: left");
        source.AppendLine(":sectnums:");
        source.AppendLine();
        for (int section = 1; section <= sectionCount; section++) {
            source.Append("[[section-").Append(section).AppendLine("]] ");
            source.Append("== Section ").Append(section).AppendLine();
            source.AppendLine();
            for (int record = 1; record <= recordsPerSection; record++) {
                int recordNumber = ((section - 1) * recordsPerSection) + record;
                source.Append("Record ").Append(recordNumber)
                    .AppendLine(": deterministic text with *bold*, _italic_, Unicode zażółć gęślą jaźń, and `code`.");
            }
            source.AppendLine();
            source.Append("* Primary item for section ").Append(section).AppendLine();
            source.Append("* Reference <<section-").Append(section).AppendLine(">>");
            source.AppendLine();
            source.AppendLine("[cols=\"1,1\"]");
            source.AppendLine("|===");
            source.AppendLine("|Key |Value");
            source.Append("|Section |").Append(section).AppendLine();
            source.AppendLine("|===");
            source.AppendLine();
        }
        return new AsciiDocBenchmarkFixture(scale, source.ToString(), sectionCount, sectionCount * recordsPerSection);
    }
}
