namespace OfficeIMO.Rtf.Benchmarks.Comparisons;

using System.Text.Json.Serialization;

internal sealed record RtfHtmlComparisonReport(
    string Scale,
    int InputBytes,
    RtfHtmlOutputEvidence OfficeIMO,
    RtfHtmlOutputEvidence RtfPipe);

internal sealed record RtfHtmlOutputEvidence(
    string Implementation,
    int OutputBytes,
    int RecordCount,
    int TableCount,
    int CellCount,
    int ImageCount,
    int SemanticTokenCount,
    string SemanticSha256,
    IReadOnlyList<string> TableCells,
    [property: JsonIgnore] string Text);
