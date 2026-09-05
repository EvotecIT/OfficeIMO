namespace OfficeIMO.Rtf;

/// <summary>Result of appending an independent semantic RTF document into another document.</summary>
public sealed class RtfDocumentMergeResult : OfficeConversionResult<RtfDocument, RtfConversionReport> {
    internal RtfDocumentMergeResult(RtfDocument document, int appendedBlockCount, RtfConversionReport report)
        : base(document, report) {
        AppendedBlockCount = appendedBlockCount;
    }

    /// <summary>Destination document after the append operation.</summary>
    public RtfDocument Document => Value;

    /// <summary>Number of body blocks appended.</summary>
    public int AppendedBlockCount { get; }

}
