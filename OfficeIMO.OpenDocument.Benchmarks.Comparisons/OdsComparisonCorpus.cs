namespace OfficeIMO.OpenDocument.Benchmarks.Comparisons;

internal sealed record OdsComparisonScale(string Name, int Rows, int Columns);

internal static class OdsComparisonCorpus {
    internal static readonly IReadOnlyList<OdsComparisonScale> Scales = new[] {
        new OdsComparisonScale("Small", 100, 8),
        new OdsComparisonScale("Normal", 1_000, 8)
    };

    internal static OdsComparisonScale Get(string name) =>
        Scales.FirstOrDefault(scale => string.Equals(scale.Name, name, StringComparison.OrdinalIgnoreCase))
        ?? throw new ArgumentException($"Unknown scale '{name}'.", nameof(name));

    internal static string Cell(int row, int column) => $"R{row:D6}C{column:D2}";
}
