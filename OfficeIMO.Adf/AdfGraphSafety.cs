using System.IO;
using System.Threading;

namespace OfficeIMO.Adf;

/// <summary>Checks mutable node graphs before any recursive serializer, validator, or projection walks them.</summary>
internal static class AdfGraphSafety {
    internal const int MaximumNodeDepth = 64;
    internal const int MaximumNodeCount = 1000000;
    internal const int MaximumJsonDepth = 256;

    internal static AdfValidationIssue? Inspect(AdfDocument document, CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        var active = new HashSet<AdfNode>();
        var frames = new Stack<Frame>();
        frames.Push(new Frame(null, document.ContentItems, "$"));
        int count = 0;
        while (frames.Count > 0) {
            cancellationToken.ThrowIfCancellationRequested();
            Frame frame = frames.Peek();
            if (frame.Index >= frame.Children.Count) {
                frames.Pop();
                if (frame.Owner != null) active.Remove(frame.Owner);
                continue;
            }
            int index = frame.Index++;
            string path = frame.Path + ".content[" + index + "]";
            AdfNode? node = frame.Children[index];
            if (node == null) return Error("ADF_NULL_NODE", path, "ADF content cannot contain null nodes.");
            if (frames.Count > MaximumNodeDepth) return Error("ADF_CONTENT_DEPTH_EXCEEDED", path, "ADF content exceeds the maximum node depth of " + MaximumNodeDepth + ".");
            if (active.Contains(node)) return Error("ADF_CYCLIC_CONTENT", path, "ADF content cannot refer to an ancestor node.");
            if (++count > MaximumNodeCount || (long)count + node.MarkItems.Count > MaximumNodeCount)
                return Error("ADF_NODE_LIMIT_EXCEEDED", path, "ADF content exceeds the maximum combined node and mark count of " + MaximumNodeCount + ".");
            count += node.MarkItems.Count;
            for (int mark = 0; mark < node.MarkItems.Count; mark++) {
                cancellationToken.ThrowIfCancellationRequested();
                if (node.MarkItems[mark] == null) return Error("ADF_NULL_MARK", path + ".marks[" + mark + "]", "ADF marks cannot contain null values.");
            }
            if (node.ContentItems.Count == 0) continue;
            active.Add(node);
            frames.Push(new Frame(node, node.ContentItems, path));
        }
        return null;
    }

    internal static void EnsureSafe(AdfDocument document, CancellationToken cancellationToken = default) {
        AdfValidationIssue? issue = Inspect(document, cancellationToken);
        if (issue != null) throw new InvalidDataException(issue.Code + " at " + issue.Path + ": " + issue.Message);
    }
    private static AdfValidationIssue Error(string code, string path, string message) => new AdfValidationIssue(code, path, message, AdfValidationSeverity.Error);
    private sealed class Frame {
        internal Frame(AdfNode? owner, IReadOnlyList<AdfNode> children, string path) { Owner = owner; Children = children; Path = path; }
        internal AdfNode? Owner { get; }
        internal IReadOnlyList<AdfNode> Children { get; }
        internal string Path { get; }
        internal int Index { get; set; }
    }
}
