using System.Collections.Concurrent;
using System.Text;

namespace OfficeIMO.Latex.Benchmarks;

internal sealed record LatexBenchmarkFixture(string Scale, string Source, int SectionCount, int RecordCount);

internal static class LatexBenchmarkCorpus {
    private static readonly ConcurrentDictionary<string, LatexBenchmarkFixture> Fixtures =
        new(StringComparer.OrdinalIgnoreCase);

    internal static IReadOnlyList<string> Scales { get; } = ["Small", "Normal", "Large"];

    internal static LatexBenchmarkFixture Get(string scale) {
        string selected = Scales.FirstOrDefault(value => string.Equals(value, scale, StringComparison.OrdinalIgnoreCase))
            ?? throw new ArgumentException("Unknown LaTeX benchmark scale: " + scale, nameof(scale));
        return Fixtures.GetOrAdd(selected, Create);
    }

    private static LatexBenchmarkFixture Create(string scale) {
        int sectionCount = scale switch {
            "Small" => 2,
            "Normal" => 100,
            _ => 1_000
        };
        const int recordsPerSection = 4;
        var source = new StringBuilder(sectionCount * 600);
        source.AppendLine("\\documentclass{article}");
        source.AppendLine("\\usepackage[utf8]{inputenc}");
        source.AppendLine("\\title{OfficeIMO deterministic LaTeX workload}");
        source.AppendLine("\\begin{document}");
        source.AppendLine("\\maketitle");
        for (int section = 1; section <= sectionCount; section++) {
            source.Append("\\section{Section ").Append(section).AppendLine("}");
            source.Append("\\label{sec:").Append(section).AppendLine("}");
            for (int record = 1; record <= recordsPerSection; record++) {
                int recordNumber = ((section - 1) * recordsPerSection) + record;
                source.Append("Record ").Append(recordNumber)
                    .Append(": deterministic text with \\textbf{bold}, \\textit{italic}, Unicode zażółć gęślą jaźń, and math $x_")
                    .Append(recordNumber).Append("^2 + y_").Append(recordNumber).AppendLine("^2$. \\\\");
            }
            source.AppendLine("\\begin{itemize}");
            source.Append("\\item Primary item for section ").Append(section).AppendLine(".");
            source.Append("\\item Reference \\ref{sec:").Append(section).AppendLine("}.");
            source.AppendLine("\\end{itemize}");
        }
        source.AppendLine("\\end{document}");
        return new LatexBenchmarkFixture(scale, source.ToString(), sectionCount, sectionCount * recordsPerSection);
    }
}
