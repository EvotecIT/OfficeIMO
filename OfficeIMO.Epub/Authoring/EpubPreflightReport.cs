namespace OfficeIMO.Epub;

/// <summary>Outcome of a publication preflight check; unperformed checks never imply success.</summary>
public enum EpubPreflightStatus {
    /// <summary>The bounded check ran without error findings.</summary>
    Passed,
    /// <summary>The check found an error or could not inspect required content.</summary>
    Failed,
    /// <summary>This check requires external validation or human review.</summary>
    NotChecked
}

/// <summary>Results for one named scope of publication review.</summary>
public sealed class EpubPreflightCheck {
    internal EpubPreflightCheck(string code, EpubPreflightStatus status, IEnumerable<EpubDiagnostic> diagnostics) {
        Code = code;
        Status = status;
        Diagnostics = Array.AsReadOnly(diagnostics.ToArray());
    }
    /// <summary>Stable check identifier.</summary>
    public string Code { get; }
    /// <summary>Whether this bounded check passed, failed, or was not performed.</summary>
    public EpubPreflightStatus Status { get; }
    /// <summary>Findings with package paths and suggested actions where available.</summary>
    public IReadOnlyList<EpubDiagnostic> Diagnostics { get; }
}

/// <summary>A snapshot of bounded native checks and outstanding independent review. Not a conformance certificate.</summary>
public sealed class EpubPreflightReport {
    internal EpubPreflightReport(IEnumerable<EpubPreflightCheck> checks) => Checks = Array.AsReadOnly(checks.ToArray());
    /// <summary>Checks in stable order, including explicit external-review gaps.</summary>
    public IReadOnlyList<EpubPreflightCheck> Checks { get; }
    /// <summary>Whether an executed native check failed; false does not establish publication readiness.</summary>
    public bool HasErrors => Checks.Any(check => check.Status == EpubPreflightStatus.Failed);
    /// <summary>Whether any review scope remains unperformed.</summary>
    public bool HasUncheckedItems => Checks.Any(check => check.Status == EpubPreflightStatus.NotChecked);
}
