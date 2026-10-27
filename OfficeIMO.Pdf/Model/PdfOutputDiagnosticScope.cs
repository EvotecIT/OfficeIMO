using System.Threading;

namespace OfficeIMO.Pdf;

/// <summary>Captures native-output diagnostics across root and nested page options for one serialization.</summary>
internal sealed class PdfOutputDiagnosticScope : IDisposable {
    private static readonly AsyncLocal<PdfOutputDiagnosticScope?> Active = new();
    private readonly PdfOutputDiagnosticScope? _previous;
    private readonly HashSet<string> _keys = new(StringComparer.Ordinal);

    internal PdfOutputDiagnosticScope() {
        _previous = Active.Value;
        Active.Value = this;
    }

    internal PdfConversionReport Report { get; } = new();
    internal static PdfConversionReport? CurrentReport => Active.Value?.Report;

    internal static bool Admit(string key, ref HashSet<string>? fallback) =>
        Active.Value is { } scope ? scope._keys.Add(key) : (fallback ??= new HashSet<string>()).Add(key);

    internal static void Capture(PdfConversionReport source, PdfConversionWarning warning) {
        if (Active.Value is { } scope && !ReferenceEquals(source, scope.Report)) scope.Report.Add(warning);
    }

    public void Dispose() => Active.Value = _previous;
}
