namespace OfficeIMO.Rtf.Benchmarks.Comparisons;

internal sealed record RtfHtmlComparisonReport(
    string Scale,
    int InputBytes,
    RtfHtmlOutputEvidence OfficeIMO,
    RtfHtmlOutputEvidence RtfPipe);

internal sealed record RtfHtmlOutputEvidence(
    string Implementation,
    int OutputBytes,
    string Text,
    int RecordCount,
    int TableCount,
    int CellCount,
    int ImageCount);
