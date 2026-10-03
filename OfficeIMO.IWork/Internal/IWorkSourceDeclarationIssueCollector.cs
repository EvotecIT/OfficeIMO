namespace OfficeIMO.IWork.Internal;

/// <summary>Retains bounded evidence for selected declarations without inspecting undecodable content.</summary>
internal sealed class IWorkSourceDeclarationIssueCollector(IWorkSourceDocument source) {
    private readonly List<IWorkSourceDeclarationIssue> _issues = new();
    private readonly HashSet<(ulong Owner, string Path)> _recorded = new();

    internal IReadOnlyList<IWorkSourceDeclarationIssue> Issues => _issues;

    internal void Record(IWorkArchiveRecord owner, string path, int? declaredValueCount,
        IWorkSourceDeclarationIssueKind kind = IWorkSourceDeclarationIssueKind.MalformedMessage) {
        source.CancellationToken.ThrowIfCancellationRequested();
        var key = (owner.Identifier, path);
        if (_recorded.Contains(key)) return;
        if (_issues.Count >= source.Options.MaximumSourceDeclarationIssues)
            throw new InvalidDataException($"iWork source declaration issues exceed the configured limit of {source.Options.MaximumSourceDeclarationIssues}.");
        _recorded.Add(key);
        _issues.Add(new IWorkSourceDeclarationIssue(new IWorkObjectIdentity(owner), path,
            declaredValueCount, kind));
    }
}
