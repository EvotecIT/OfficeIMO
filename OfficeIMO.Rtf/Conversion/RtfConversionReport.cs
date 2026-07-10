using OfficeIMO.Rtf.Diagnostics;

namespace OfficeIMO.Rtf;

/// <summary>Shared fidelity and policy report used by every RTF conversion adapter.</summary>
public sealed class RtfConversionReport {
    private readonly List<RtfConversionDiagnostic> _diagnostics = new List<RtfConversionDiagnostic>();

    /// <summary>Snapshot of recorded diagnostics in conversion order.</summary>
    public IReadOnlyList<RtfConversionDiagnostic> Diagnostics => _diagnostics.AsReadOnly();

    /// <summary>Whether the report contains a flattened, omitted, blocked, or error condition.</summary>
    public bool HasLoss => _diagnostics.Any(IsLoss);

    /// <summary>Adds a conversion diagnostic.</summary>
    public RtfConversionDiagnostic Add(
        RtfConversionSeverity severity,
        string code,
        string message,
        RtfConversionAction action,
        string? sourcePath = null,
        string? feature = null,
        int count = 1,
        string? detail = null) {
        var diagnostic = new RtfConversionDiagnostic(severity, code, message, action, sourcePath, feature, count, detail);
        _diagnostics.Add(diagnostic);
        return diagnostic;
    }

    /// <summary>Adds an existing conversion diagnostic.</summary>
    public void Add(RtfConversionDiagnostic diagnostic) {
        _diagnostics.Add(diagnostic ?? throw new ArgumentNullException(nameof(diagnostic)));
    }

    /// <summary>Adds diagnostics from another report.</summary>
    public void Merge(RtfConversionReport? report) {
        if (report == null || ReferenceEquals(report, this)) return;
        foreach (RtfConversionDiagnostic diagnostic in report.Diagnostics) {
            _diagnostics.Add(diagnostic);
        }
    }

    /// <summary>Clears diagnostics from a previous conversion run.</summary>
    public void Clear() => _diagnostics.Clear();

    /// <summary>Maps parser and binder diagnostics into the shared report.</summary>
    public void AddReadDiagnostics(IEnumerable<RtfDiagnostic>? diagnostics, string? sourcePath = null) {
        if (diagnostics == null) return;
        foreach (RtfDiagnostic diagnostic in diagnostics) {
            RtfConversionSeverity severity = diagnostic.Severity == RtfDiagnosticSeverity.Error
                ? RtfConversionSeverity.Error
                : diagnostic.Severity == RtfDiagnosticSeverity.Warning
                    ? RtfConversionSeverity.Warning
                    : RtfConversionSeverity.Information;
            RtfConversionAction action = diagnostic.Code == "RTF105" || diagnostic.Code == "RTF106" || diagnostic.Code == "RTF107"
                ? RtfConversionAction.Blocked
                : RtfConversionAction.Omitted;
            Add(severity, diagnostic.Code, diagnostic.Message, action, sourcePath, detail: diagnostic.Position.ToString(CultureInfo.InvariantCulture));
        }
    }

    /// <summary>Throws when the report contains any fidelity loss or blocked feature.</summary>
    public void RequireNoLoss() {
        if (HasLoss) throw new RtfConversionLossException(this);
    }

    private static bool IsLoss(RtfConversionDiagnostic diagnostic) =>
        diagnostic.Severity == RtfConversionSeverity.Error
        || diagnostic.Action == RtfConversionAction.Flattened
        || diagnostic.Action == RtfConversionAction.Omitted
        || diagnostic.Action == RtfConversionAction.Blocked;
}
