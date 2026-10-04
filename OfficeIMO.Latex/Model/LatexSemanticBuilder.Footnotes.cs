using System.Threading;

namespace OfficeIMO.Latex;

internal static partial class LatexSemanticBuilder {
    internal static bool IsInsideFootnote(LatexSyntaxNode node) {
        for (LatexSyntaxNode? parent = node.Parent; parent != null; parent = parent.Parent)
            if (parent.Kind == LatexSyntaxKind.Command && parent.Value == "footnote") return true;
        return false;
    }

    private static IReadOnlyList<LatexFootnote> BuildFootnotes(
        IReadOnlyList<LatexCommand> commands, CancellationToken cancellationToken) {
        var result = new List<LatexFootnote>();
        foreach (LatexCommand command in commands) {
            cancellationToken.ThrowIfCancellationRequested();
            if (command.Name == "footnote" && command.GetRequiredArgument(0) != null)
                result.Add(new LatexFootnote(command));
        }
        return result.Count == 0 ? Array.Empty<LatexFootnote>() : result.ToArray();
    }
}
