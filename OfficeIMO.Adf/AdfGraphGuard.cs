using System.IO;
using System.Text;
using System.Text.Json;
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

    internal static void Check(AdfDocument document, AdfProcessingOptions options) {
        options.Check();
        if (document.ContentItems.Count > options.MaxNodes) throw Limit("MaxNodes");
        var ancestors = new HashSet<AdfNode>();
        var pending = new Stack<(AdfNode? Node, int Depth, bool Exit)>();
        foreach (AdfNode node in document.ContentItems) pending.Push((node, 1, false));
        long nodes = 0;
        long text = JsonTextLength(document.ExtensionItems.Values, options);
        while (pending.Count > 0) {
            options.CancellationToken.ThrowIfCancellationRequested();
            var frame = pending.Pop();
            if (frame.Node == null) {
                if (++nodes > options.MaxNodes) throw Limit("MaxNodes");
                continue; // Structural validation reports null content.
            }
            if (frame.Exit) { ancestors.Remove(frame.Node); continue; }
            if (frame.Depth > options.MaxDepth) throw Limit("MaxDepth");
            nodes += 1L + frame.Node.MarkItems.Count;
            if (nodes > options.MaxNodes) throw Limit("MaxNodes");
            text += frame.Node.Text?.Length ?? 0;
            if (frame.Node.AttributeItems.Count > 0) text += JsonTextLength(frame.Node.AttributeItems.Values, options);
            if (frame.Node.ExtensionItems.Count > 0) text += JsonTextLength(frame.Node.ExtensionItems.Values, options);
            foreach (AdfMark mark in frame.Node.MarkItems) {
                if (mark != null && mark.AttributeItems.Count > 0) text += JsonTextLength(mark.AttributeItems.Values, options);
                if (mark != null && mark.ExtensionItems.Count > 0) text += JsonTextLength(mark.ExtensionItems.Values, options);
            }
            if (text > options.MaxTextCharacters) throw Limit("MaxTextCharacters");
            if (!ancestors.Add(frame.Node)) throw new InvalidOperationException("ADF content contains a cycle.");
            if (frame.Node.ContentItems.Count > options.MaxNodes - nodes) throw Limit("MaxNodes");
            pending.Push((frame.Node, frame.Depth, true));
            for (int i = frame.Node.ContentItems.Count - 1; i >= 0; i--) pending.Push((frame.Node.ContentItems[i], frame.Depth + 1, false));
        }
        if (text > options.MaxTextCharacters) throw Limit("MaxTextCharacters");
    }

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

    private static long JsonTextLength(IEnumerable<JsonElement> values, AdfProcessingOptions options) {
        var pending = new Stack<JsonElement>(values);
        long length = 0;
        while (pending.Count > 0) {
            options.CancellationToken.ThrowIfCancellationRequested();
            JsonElement value = pending.Pop();
            if (value.ValueKind == JsonValueKind.String) length += value.GetString()?.Length ?? 0;
            else if (value.ValueKind == JsonValueKind.Object) foreach (JsonProperty property in value.EnumerateObject()) pending.Push(property.Value);
            else if (value.ValueKind == JsonValueKind.Array) foreach (JsonElement item in value.EnumerateArray()) pending.Push(item);
            if (length > options.MaxTextCharacters) throw Limit("MaxTextCharacters");
        }
        return length;
    }

    internal static InvalidDataException Limit(string name) => new InvalidDataException("ADF operation exceeded " + name + ".");
}
