namespace OfficeIMO.Pdf;

/// <summary>
/// Common high-level PDF export profiles that converters can map to their own detailed options.
/// </summary>
public enum PdfExportProfile {
    /// <summary>Preserve authored document content and visual features where the converter supports them.</summary>
    Faithful = 0,

    /// <summary>Prefer smaller, simpler output by omitting heavyweight visual content where supported.</summary>
    Lightweight = 1,

    /// <summary>Prefer page setup, repeated headers/footers, and pagination choices intended for printing.</summary>
    PrintReady = 2,

    /// <summary>Prefer text and table content over decorative media, charts, and backgrounds where supported.</summary>
    TextOnly = 3
}
