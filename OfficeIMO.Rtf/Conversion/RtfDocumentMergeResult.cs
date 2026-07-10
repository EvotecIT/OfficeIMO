namespace OfficeIMO.Rtf;

/// <summary>Result of appending an independent semantic RTF document into another document.</summary>
public sealed class RtfDocumentMergeResult {
    internal RtfDocumentMergeResult(RtfDocument document, int appendedBlockCount, RtfConversionReport report) {
        Document = document;
        AppendedBlockCount = appendedBlockCount;
        Report = report;
    }

    /// <summary>Destination document after the append operation.</summary>
    public RtfDocument Document { get; }

    /// <summary>Number of body blocks appended.</summary>
    public int AppendedBlockCount { get; }

    /// <summary>Resource substitutions and semantic degradation produced by the merge.</summary>
    public RtfConversionReport Report { get; }
}
