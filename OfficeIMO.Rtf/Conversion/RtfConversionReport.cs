using OfficeIMO.Rtf.Diagnostics;

namespace OfficeIMO.Rtf;

/// <summary>Shared fidelity and policy report used by every RTF conversion adapter.</summary>
public sealed class RtfConversionReport : OfficeConversionReport {
    private readonly List<RtfConversionDiagnostic> _diagnostics = new List<RtfConversionDiagnostic>();

    /// <summary>Snapshot of recorded diagnostics in conversion order.</summary>
    public IReadOnlyList<RtfConversionDiagnostic> NativeDiagnostics => _diagnostics.AsReadOnly();

    /// <summary>Category-preserving diagnostics for composed conversion routes.</summary>
    protected override IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> GetDiagnostics() => FidelityDiagnostics;
    private IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> FidelityDiagnostics => Array.AsReadOnly(
        _diagnostics.Select(static diagnostic => new global::OfficeIMO.OfficeDiagnostic(
            string.IsNullOrWhiteSpace(diagnostic.Code) ? "RTF_DIAGNOSTIC_UNSPECIFIED" : diagnostic.Code,
            global::OfficeIMO.OfficeDiagnosticSeverity.Warning,
            diagnostic.Message,
            diagnostic.SourcePath,
            "OfficeIMO.Rtf",
            GetLossKind(diagnostic))).ToArray());

    /// <summary>Whether the report contains a flattened, omitted, blocked, or error condition.</summary>
    public override bool HasLoss => _diagnostics.Any(static diagnostic => GetLossKind(diagnostic) != OfficeConversionLossKind.None);

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
        foreach (RtfConversionDiagnostic diagnostic in report.NativeDiagnostics) {
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
            RtfConversionAction action = GetReadAction(diagnostic);
            Add(severity, diagnostic.Code, diagnostic.Message, action, sourcePath, detail: diagnostic.Position.ToString(CultureInfo.InvariantCulture));
        }
    }

    /// <summary>Throws when the report contains any fidelity loss or blocked feature.</summary>
    protected override OfficeConversionException CreateLossException() => new RtfConversionLossException(this);

    private static OfficeConversionLossKind GetLossKind(RtfConversionDiagnostic diagnostic) =>
        diagnostic.Severity == RtfConversionSeverity.Error || diagnostic.Action == RtfConversionAction.Blocked
            ? OfficeConversionLossKind.Failure
            : diagnostic.Action == RtfConversionAction.Omitted
                ? OfficeConversionLossKind.Omission
                : diagnostic.Action == RtfConversionAction.Flattened
                    ? OfficeConversionLossKind.Approximation
                    : OfficeConversionLossKind.None;

    private static RtfConversionAction GetReadAction(RtfDiagnostic diagnostic) {
        if (diagnostic.Severity == RtfDiagnosticSeverity.Error ||
            diagnostic.Code is "RTF105" or "RTF106" or "RTF107") {
            return RtfConversionAction.Blocked;
        }
        return diagnostic.Code switch {
            "RTF002" or "RTF012" or "RTF103" => RtfConversionAction.Flattened,
            "RTF010" or "RTF011" or "RTF014" or "RTF101" or "RTF102" => RtfConversionAction.Omitted,
            _ => diagnostic.Severity == RtfDiagnosticSeverity.Warning
                ? RtfConversionAction.Flattened
                : RtfConversionAction.Preserved
        };
    }
}
