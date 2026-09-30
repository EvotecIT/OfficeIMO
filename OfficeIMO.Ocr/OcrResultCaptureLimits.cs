using System;

namespace OfficeIMO.Ocr;

/// <summary>Bounds provider-owned collections copied during a recognition invocation.</summary>
public sealed class OcrResultCaptureLimits {
    /// <summary>Creates immutable retention limits. Omitted counts remain available on the captured result.</summary>
    public OcrResultCaptureLimits(int maxSpans = int.MaxValue, int maxDiagnostics = int.MaxValue,
        int maxDiagnosticAttributes = int.MaxValue) {
        if (maxSpans < 0) throw new ArgumentOutOfRangeException(nameof(maxSpans));
        if (maxDiagnostics < 0) throw new ArgumentOutOfRangeException(nameof(maxDiagnostics));
        if (maxDiagnosticAttributes < 0) throw new ArgumentOutOfRangeException(nameof(maxDiagnosticAttributes));
        MaxSpans = maxSpans; MaxDiagnostics = maxDiagnostics; MaxDiagnosticAttributes = maxDiagnosticAttributes;
    }
    /// <summary>Maximum retained text spans.</summary>
    public int MaxSpans { get; }
    /// <summary>Maximum retained diagnostics. Terminal errors are checked even outside this prefix.</summary>
    public int MaxDiagnostics { get; }
    /// <summary>Maximum retained attributes across all retained diagnostics.</summary>
    public int MaxDiagnosticAttributes { get; }
}
