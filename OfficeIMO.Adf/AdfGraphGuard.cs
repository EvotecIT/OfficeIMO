using System.IO;
using System.Text;
using OfficeIMO.Markdown;

namespace OfficeIMO.Adf;

internal static class AdfGraphGuard {
    internal static void Input(string text, AdfProcessingOptions options) {
        options.Check();
        if (Encoding.UTF8.GetByteCount(text) > options.MaxInputBytes) throw Limit("MaxInputBytes");
    }

    internal static void Output(string text, AdfProcessingOptions options) {
        options.CancellationToken.ThrowIfCancellationRequested();
        if (text.Length > options.MaxOutputCharacters) throw Limit("MaxOutputCharacters");
    }

    internal static void Check(AdfDocument document, AdfProcessingOptions options) => AdfGraphSafety.EnsureSafe(document, options);

    internal static void CheckMarkdown(MarkdownDoc document, AdfProcessingOptions options) {
        options.Check();
        var ancestors = new HashSet<MarkdownObject>(MarkdownReferenceComparer.Instance);
        var pending = new Stack<(MarkdownObject Node, int Depth, bool Exit)>();
        pending.Push((document, 0, false));
        long nodes = 0;
        while (pending.Count > 0) {
            options.CancellationToken.ThrowIfCancellationRequested();
            var frame = pending.Pop();
            if (frame.Exit) { ancestors.Remove(frame.Node); continue; }
            if (frame.Depth > options.MaxDepth) throw Limit("MaxDepth");
            if (frame.Depth > 0 && ++nodes > options.MaxNodes) throw Limit("MaxNodes");
            if (!ancestors.Add(frame.Node)) throw new InvalidOperationException("Markdown content contains a cycle.");
            IReadOnlyList<MarkdownObject> children = frame.Node.ChildObjects;
            if (children.Count > options.MaxNodes - nodes) throw Limit("MaxNodes");
            pending.Push((frame.Node, frame.Depth, true));
            for (int i = children.Count - 1; i >= 0; i--) pending.Push((children[i], frame.Depth + 1, false));
        }
    }

    private sealed class MarkdownReferenceComparer : IEqualityComparer<MarkdownObject> {
        internal static readonly MarkdownReferenceComparer Instance = new MarkdownReferenceComparer();
        public bool Equals(MarkdownObject? x, MarkdownObject? y) => ReferenceEquals(x, y);
        public int GetHashCode(MarkdownObject value) => System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(value);
    }

    internal static InvalidDataException Limit(string name) => new InvalidDataException("ADF operation exceeded " + name + ".");
}
