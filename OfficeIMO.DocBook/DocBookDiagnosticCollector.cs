using System.Collections.Generic;

namespace OfficeIMO.DocBook;

/// <summary>Keeps repeated diagnostics bounded while retaining the format's diagnostic contract.</summary>
internal sealed class DocBookDiagnosticCollector : OfficeDiagnosticCollector<DocBookDiagnostic> {
    internal DocBookDiagnosticCollector(int maxDetailedPerCode) : base(maxDetailedPerCode) { }

    internal void Add(DocBookDiagnostic diagnostic) => Collect(diagnostic);
    internal IReadOnlyList<DocBookDiagnostic> ToArray() => GetDiagnostics();

    protected override string CodeOf(DocBookDiagnostic diagnostic) => diagnostic.Code;
    protected override int SeverityOf(DocBookDiagnostic diagnostic) => (int)diagnostic.Severity;
    protected override OfficeConversionLossKind LossKindOf(DocBookDiagnostic diagnostic) => diagnostic.LossKind;
    protected override DocBookDiagnostic CreateSummary(string code, int severity, string message, OfficeConversionLossKind lossKind) =>
        new DocBookDiagnostic(code, (DocBookDiagnosticSeverity)severity, message, path: null, lossKind: lossKind);
}
