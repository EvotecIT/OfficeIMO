using System.Collections.Concurrent;
using System.Text;

namespace OfficeIMO.Markup.Benchmarks;

internal sealed record OfficeMarkupBenchmarkFixture(string Scale, string Source, int SectionCount, int RecordCount);

internal static class OfficeMarkupBenchmarkCorpus {
    private static readonly ConcurrentDictionary<string, OfficeMarkupBenchmarkFixture> Fixtures = new(StringComparer.OrdinalIgnoreCase);
    internal static IReadOnlyList<string> Scales { get; } = ["Small", "Normal", "Large"];

    internal static OfficeMarkupBenchmarkFixture Get(string scale) {
        string selected = Scales.FirstOrDefault(value => string.Equals(value, scale, StringComparison.OrdinalIgnoreCase))
            ?? throw new ArgumentException("Unknown Office Markup benchmark scale: " + scale, nameof(scale));
        return Fixtures.GetOrAdd(selected, Create);
    }

    private static OfficeMarkupBenchmarkFixture Create(string scale) {
        int sectionCount = scale switch { "Small" => 2, "Normal" => 100, _ => 1_000 };
        const int recordsPerSection = 4;
        var source = new StringBuilder(sectionCount * 600);
        source.AppendLine("# OfficeIMO deterministic semantic-markup workload");
        source.AppendLine();
        for (int section = 1; section <= sectionCount; section++) {
            source.Append("## Section ").Append(section).AppendLine();
            source.AppendLine();
            for (int record = 1; record <= recordsPerSection; record++) {
                int number = ((section - 1) * recordsPerSection) + record;
                source.Append("Record ").Append(number)
                    .AppendLine(" has **strong**, _emphasis_, `code`, and Unicode zażółć gęślą jaźń.");
            }
            source.AppendLine();
            source.Append("- Primary item ").Append(section).AppendLine();
            source.Append("- Secondary item ").Append(section).AppendLine();
            source.AppendLine();
            source.AppendLine("| Key | Value |");
            source.AppendLine("| --- | --- |");
            source.Append("| Section | ").Append(section).AppendLine(" |");
            source.Append("| Status | Ready ").Append(section).AppendLine(" |");
            source.AppendLine();
        }
        return new OfficeMarkupBenchmarkFixture(scale, source.ToString(), sectionCount, sectionCount * recordsPerSection);
    }
}
