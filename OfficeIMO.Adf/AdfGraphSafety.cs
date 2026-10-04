using System.IO;
using System.Threading;
using System.Text.Json;

namespace OfficeIMO.Adf;

/// <summary>Checks mutable node graphs before any recursive serializer, validator, or projection walks them.</summary>
internal static class AdfGraphSafety {
    internal static AdfValidationIssue? Inspect(AdfDocument document, AdfProcessingOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        options ??= new AdfProcessingOptions();
        options.Check();
        CancellationToken cancellationToken = options.CancellationToken;
        var active = new HashSet<AdfNode>();
        var frames = new Stack<Frame>();
        frames.Push(new Frame(null, document.ContentItems, "$"));
        long count = 0;
        long text = JsonTextLength(document.ExtensionItems.Values, options);
        if (text > options.MaxTextCharacters) return Error("ADF_TEXT_LIMIT_EXCEEDED", "$", "ADF operation exceeded MaxTextCharacters.");
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
            if (frames.Count > options.MaxDepth) return Error("ADF_CONTENT_DEPTH_EXCEEDED", path, "ADF operation exceeded MaxDepth.");
            if (active.Contains(node)) return Error("ADF_CYCLIC_CONTENT", path, "ADF content cannot refer to an ancestor node.");
            if (++count > options.MaxNodes || count + node.MarkItems.Count > options.MaxNodes)
                return Error("ADF_NODE_LIMIT_EXCEEDED", path, "ADF operation exceeded MaxNodes.");
            count += node.MarkItems.Count;
            text += node.Text?.Length ?? 0;
            text += JsonTextLength(node.AttributeItems.Values, options) + JsonTextLength(node.ExtensionItems.Values, options);
            for (int mark = 0; mark < node.MarkItems.Count; mark++) {
                cancellationToken.ThrowIfCancellationRequested();
                if (node.MarkItems[mark] == null) return Error("ADF_NULL_MARK", path + ".marks[" + mark + "]", "ADF marks cannot contain null values.");
                text += JsonTextLength(node.MarkItems[mark].AttributeItems.Values, options) + JsonTextLength(node.MarkItems[mark].ExtensionItems.Values, options);
            }
            if (text > options.MaxTextCharacters) return Error("ADF_TEXT_LIMIT_EXCEEDED", path, "ADF operation exceeded MaxTextCharacters.");
            if (node.ContentItems.Count == 0) continue;
            active.Add(node);
            frames.Push(new Frame(node, node.ContentItems, path));
        }
        return null;
    }

    internal static void EnsureSafe(AdfDocument document, AdfProcessingOptions? options = null) {
        AdfValidationIssue? issue = Inspect(document, options);
        if (issue != null) throw new InvalidDataException(issue.Code + " at " + issue.Path + ": " + issue.Message);
    }
    private static long JsonTextLength(IEnumerable<JsonElement> values, AdfProcessingOptions options) {
        var pending = new Stack<JsonElement>(values);
        long length = 0;
        while (pending.Count > 0) {
            options.CancellationToken.ThrowIfCancellationRequested();
            JsonElement value = pending.Pop();
            if (value.ValueKind == JsonValueKind.String) length += value.GetString()?.Length ?? 0;
            else if (value.ValueKind == JsonValueKind.Object) foreach (JsonProperty property in value.EnumerateObject()) pending.Push(property.Value);
            else if (value.ValueKind == JsonValueKind.Array) foreach (JsonElement item in value.EnumerateArray()) pending.Push(item);
            if (length > options.MaxTextCharacters) break;
        }
        return length;
    }
    private static AdfValidationIssue Error(string code, string path, string message) => new AdfValidationIssue(code, path, message, AdfValidationSeverity.Error, isGraphSafetyFailure: true);
    private sealed class Frame {
        internal Frame(AdfNode? owner, IReadOnlyList<AdfNode> children, string path) { Owner = owner; Children = children; Path = path; }
        internal AdfNode? Owner { get; }
        internal IReadOnlyList<AdfNode> Children { get; }
        internal string Path { get; }
        internal int Index { get; set; }
    }
}
