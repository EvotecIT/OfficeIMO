using System;

namespace OfficeIMO;

/// <summary>Named cross-format output intent.</summary>
public enum OfficeConversionProfile {
    /// <summary>Preserve authored content and visual features where supported.</summary>
    Faithful,
    /// <summary>Prefer smaller output while retaining useful document structure.</summary>
    Lightweight,
    /// <summary>Prefer page setup and output intended for printing.</summary>
    PrintReady,
    /// <summary>Prefer text and tables over decorative visual content.</summary>
    TextOnly
}
