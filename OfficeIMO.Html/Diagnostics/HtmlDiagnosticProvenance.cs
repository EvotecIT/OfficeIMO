namespace OfficeIMO.Html;

/// <summary>Source-to-target provenance attached to every shared HTML diagnostic.</summary>
public sealed class HtmlDiagnosticProvenance {
    internal HtmlDiagnosticProvenance(string sourceAddress, int sourceLine, int sourceColumn, string targetAddress) {
        SourceAddress = sourceAddress;
        SourceLine = sourceLine;
        SourceColumn = sourceColumn;
        TargetAddress = targetAddress;
    }

    /// <summary>HTML selector, URI, or stable source scope associated with the diagnostic.</summary>
    public string SourceAddress { get; }

    /// <summary>One-based HTML source line, or zero when unavailable.</summary>
    public int SourceLine { get; }

    /// <summary>One-based HTML source column, or zero when unavailable.</summary>
    public int SourceColumn { get; }

    /// <summary>Target address such as a paragraph, cell, slide/shape, page/element, or renderer scope.</summary>
    public string TargetAddress { get; }
}
