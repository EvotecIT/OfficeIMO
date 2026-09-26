using System.Collections.ObjectModel;

namespace OfficeIMO.Project;

/// <summary>Diagnostic importance for project validation and fidelity.</summary>
public enum ProjectDiagnosticSeverity {
    /// <summary>Preserved or contextual information.</summary>
    Information,
    /// <summary>A limitation or risk requiring caller attention.</summary>
    Warning,
    /// <summary>An error that prevents safe serialization.</summary>
    Error
}

/// <summary>A stable code and source/object location describing validation or preservation behavior.</summary>
public sealed class ProjectDiagnostic {
    internal ProjectDiagnostic(string code, ProjectDiagnosticSeverity severity, string message, string location, bool loss = false)
        : this(code, severity, message, location,
            loss ? OfficeConversionLossKind.Approximation : OfficeConversionLossKind.None) {
    }
    internal ProjectDiagnostic(string code, ProjectDiagnosticSeverity severity, string message, string location,
        OfficeConversionLossKind lossKind) {
        if (!Enum.IsDefined(typeof(OfficeConversionLossKind), lossKind)) throw new ArgumentOutOfRangeException(nameof(lossKind));
        Code = code;
        Severity = severity;
        Message = message;
        Location = location;
        LossKind = severity == ProjectDiagnosticSeverity.Error ? OfficeConversionLossKind.Failure : lossKind;
    }
    /// <summary>Stable machine-readable diagnostic code.</summary>
    public string Code { get; }
    /// <summary>Importance of the finding.</summary>
    public ProjectDiagnosticSeverity Severity { get; }
    /// <summary>Actionable explanation.</summary>
    public string Message { get; }
    /// <summary>Entity UID or XML location.</summary>
    public string Location { get; }
    /// <summary>Exact fidelity category retained across composed project routes.</summary>
    public OfficeConversionLossKind LossKind { get; }
    /// <summary>True for known or possible semantic loss, even if raw content is retained.</summary>
    public bool RepresentsLoss => LossKind != OfficeConversionLossKind.None;
}

/// <summary>A snapshot of validation/fidelity findings for one model revision.</summary>
public sealed class ProjectReport : IOfficeConversionReport {
    internal ProjectReport(long revision, IEnumerable<ProjectDiagnostic> diagnostics) {
        ModelRevision = revision; Diagnostics = new ReadOnlyCollection<ProjectDiagnostic>(diagnostics.ToArray());
        FidelityDiagnostics = new ReadOnlyCollection<OfficeConversionFidelityDiagnostic>(Diagnostics.Select(static diagnostic =>
            new OfficeConversionFidelityDiagnostic(
                diagnostic.Code,
                diagnostic.Message,
                diagnostic.LossKind,
                "OfficeIMO.Project",
                diagnostic.Location)).ToArray());
    }
    /// <summary>The model revision assessed, not a promise about future mutations.</summary>
    public long ModelRevision { get; }
    /// <summary>Immutable findings.</summary>
    public IReadOnlyList<ProjectDiagnostic> Diagnostics { get; }
    /// <summary>Category-preserving project diagnostics.</summary>
    public IReadOnlyList<OfficeConversionFidelityDiagnostic> FidelityDiagnostics { get; }
    /// <summary>True when any finding prevents save.</summary>
    public bool HasErrors => Diagnostics.Any(d => d.Severity == ProjectDiagnosticSeverity.Error);
    /// <inheritdoc />
    public bool HasLoss => FidelityDiagnostics.Any(d => d.LossKind != OfficeConversionLossKind.None);
    /// <summary>Throws with the first validation error when the model is inconsistent.</summary>
    public void ThrowIfErrors() {
        var error = Diagnostics.FirstOrDefault(d => d.Severity == ProjectDiagnosticSeverity.Error);
        if (error != null) throw new InvalidDataException(error.Code + " at " + error.Location + ": " + error.Message);
    }
    /// <inheritdoc />
    public void RequireNoLoss() {
        ThrowIfErrors();
        var loss = Diagnostics.FirstOrDefault(d => d.RepresentsLoss);
        if (loss != null) throw new InvalidOperationException(loss.Code + " at " + loss.Location + ": " + loss.Message);
    }
}
