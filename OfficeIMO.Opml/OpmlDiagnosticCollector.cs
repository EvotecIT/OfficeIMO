using System.Collections.Generic;

namespace OfficeIMO.Opml;

/// <summary>Keeps repeated diagnostics bounded while retaining the format's diagnostic contract.</summary>
internal sealed class OpmlDiagnosticCollector : OfficeDiagnosticCollector<OpmlDiagnostic> {
    internal OpmlDiagnosticCollector(int maxDetailedPerCode) : base(maxDetailedPerCode) { }

    internal void Add(OpmlDiagnostic diagnostic) => Collect(diagnostic);
    internal IReadOnlyList<OpmlDiagnostic> ToArray() => GetDiagnostics();

    protected override string CodeOf(OpmlDiagnostic diagnostic) => diagnostic.Code;
    protected override int SeverityOf(OpmlDiagnostic diagnostic) => (int)diagnostic.Severity;
    protected override OfficeConversionLossKind LossKindOf(OpmlDiagnostic diagnostic) => diagnostic.LossKind;
    protected override OpmlDiagnostic CreateSummary(string code, int severity, string message, OfficeConversionLossKind lossKind) =>
        new OpmlDiagnostic(code, (OpmlDiagnosticSeverity)severity, message, path: null, lossKind: lossKind);
}
