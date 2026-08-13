using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Runtime.CompilerServices;

namespace OfficeIMO.Pdf;

internal static class PdfActionPayloadFingerprint {
    private const int MaximumDepth = 32;
    private const int MaximumNodes = 4096;
    private static readonly ConditionalWeakTable<Dictionary<int, PdfIndirectObject>, PageNumberLookupCache> PageNumberLookups = new();

    internal static string? Create(
        PdfDictionary action,
        Dictionary<int, PdfIndirectObject> objects,
        PdfReadLimits limits) {
        var builder = new StringBuilder();
        var activeReferences = new HashSet<(int ObjectNumber, int Generation)>();
        PageNumberLookup pageNumberLookup = PageNumberLookups.GetValue(
            objects,
            static source => new PageNumberLookupCache(source)).Get(limits);
        if (!pageNumberLookup.IsComplete) return null;
        IReadOnlyDictionary<int, int> pageNumbers = pageNumberLookup.Value;
        int nodes = 0;
        bool complete = true;
        AppendDictionary(builder, action, objects, pageNumbers, activeReferences, depth: 0, ref nodes, ref complete, isActionRoot: true);
        return complete ? builder.ToString() : null;
    }

    private static void AppendObject(
        StringBuilder builder,
        PdfObject? value,
        Dictionary<int, PdfIndirectObject> objects,
        IReadOnlyDictionary<int, int> pageNumbers,
        HashSet<(int ObjectNumber, int Generation)> activeReferences,
        int depth,
        ref int nodes,
        ref bool complete) {
        nodes++;
        if (depth > MaximumDepth || nodes > MaximumNodes) {
            complete = false;
            return;
        }

        switch (value) {
            case null:
            case PdfNull:
                builder.Append('z');
                return;
            case PdfBoolean boolean:
                builder.Append(boolean.Value ? "b1" : "b0");
                return;
            case PdfNumber number:
                builder.Append('n').Append(number.Value.ToString("R", CultureInfo.InvariantCulture));
                return;
            case PdfName name:
                AppendText(builder, 'N', name.Name);
                return;
            case PdfStringObj text:
                AppendText(builder, 'S', Convert.ToBase64String(text.RawBytes));
                return;
            case PdfReference reference:
                AppendReference(builder, reference, objects, pageNumbers, activeReferences, depth, ref nodes, ref complete);
                return;
            case PdfArray array:
                builder.Append('[');
                for (int i = 0; i < array.Items.Count; i++) {
                    AppendObject(builder, array.Items[i], objects, pageNumbers, activeReferences, depth + 1, ref nodes, ref complete);
                    builder.Append(';');
                }
                builder.Append(']');
                return;
            case PdfDictionary dictionary:
                AppendDictionary(builder, dictionary, objects, pageNumbers, activeReferences, depth + 1, ref nodes, ref complete, isActionRoot: false);
                return;
            case PdfStream stream:
                builder.Append("stream:");
                AppendDictionary(builder, stream.Dictionary, objects, pageNumbers, activeReferences, depth + 1, ref nodes, ref complete, isActionRoot: false);
#if NET8_0_OR_GREATER
                AppendText(builder, 'H', Convert.ToBase64String(SHA256.HashData(stream.Data)));
#else
                using (SHA256 sha256 = SHA256.Create()) {
                    AppendText(builder, 'H', Convert.ToBase64String(sha256.ComputeHash(stream.Data)));
                }
#endif
                return;
            default:
                AppendText(builder, '?', value.GetType().FullName ?? value.GetType().Name);
                return;
        }
    }

    private static void AppendReference(
        StringBuilder builder,
        PdfReference reference,
        Dictionary<int, PdfIndirectObject> objects,
        IReadOnlyDictionary<int, int> pageNumbers,
        HashSet<(int ObjectNumber, int Generation)> activeReferences,
        int depth,
        ref int nodes,
        ref bool complete) {
        var key = (reference.ObjectNumber, reference.Generation);
        if (!activeReferences.Add(key)) {
            builder.Append("cycle:").Append(reference.ObjectNumber).Append(':').Append(reference.Generation);
            return;
        }
        try {
            if (!PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect)) {
                builder.Append("ref:").Append(reference.ObjectNumber).Append(':').Append(reference.Generation);
                return;
            }
            if (indirect.Value is PdfDictionary dictionary &&
                string.Equals(dictionary.Get<PdfName>("Type")?.Name, "Page", StringComparison.Ordinal)) {
                if (pageNumbers.TryGetValue(reference.ObjectNumber, out int pageNumber)) {
                    builder.Append("page:").Append(pageNumber);
                } else {
                    builder.Append("page-ref:").Append(reference.ObjectNumber).Append(':').Append(reference.Generation);
                }
                return;
            }
            AppendObject(builder, indirect.Value, objects, pageNumbers, activeReferences, depth + 1, ref nodes, ref complete);
        } finally {
            activeReferences.Remove(key);
        }
    }

    private static void AppendDictionary(
        StringBuilder builder,
        PdfDictionary dictionary,
        Dictionary<int, PdfIndirectObject> objects,
        IReadOnlyDictionary<int, int> pageNumbers,
        HashSet<(int ObjectNumber, int Generation)> activeReferences,
        int depth,
        ref int nodes,
        ref bool complete,
        bool isActionRoot) {
        builder.Append('{');
        foreach (string key in dictionary.Items.Keys.OrderBy(static key => key, StringComparer.Ordinal)) {
            if (isActionRoot &&
                (string.Equals(key, "S", StringComparison.Ordinal) ||
                 string.Equals(key, "Next", StringComparison.Ordinal))) continue;
            AppendText(builder, 'K', key);
            AppendObject(builder, dictionary.Items[key], objects, pageNumbers, activeReferences, depth + 1, ref nodes, ref complete);
            builder.Append(';');
        }
        builder.Append('}');
    }

    private static PageNumberLookup BuildPageNumberLookup(Dictionary<int, PdfIndirectObject> objects, PdfReadLimits limits) {
        var result = new Dictionary<int, int>();
        PdfDictionary? catalog = PdfSyntax.FindCatalog(objects);
        if (catalog == null || !catalog.Items.TryGetValue("Pages", out PdfObject? pages)) return new PageNumberLookup(result, isComplete: true);
        var visited = new HashSet<int>();
        bool complete = true;
        AddPageTreeNode(pages, objects, visited, result, depth: 0, limits, ref complete);
        return new PageNumberLookup(result, complete);
    }

    private static void AddPageTreeNode(
        PdfObject node,
        Dictionary<int, PdfIndirectObject> objects,
        HashSet<int> visited,
        Dictionary<int, int> pageNumbers,
        int depth,
        PdfReadLimits limits,
        ref bool complete) {
        if (depth > limits.MaxPageTreeDepth) {
            complete = false;
            return;
        }
        int objectNumber = 0;
        if (node is PdfReference reference) {
            objectNumber = reference.ObjectNumber;
            if (!visited.Add(objectNumber) ||
                !PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect)) return;
            if (visited.Count > limits.MaxPageTreeNodes) {
                complete = false;
                return;
            }
            node = indirect.Value;
        }
        if (node is not PdfDictionary dictionary) return;
        if (string.Equals(dictionary.Get<PdfName>("Type")?.Name, "Page", StringComparison.Ordinal)) {
            if (objectNumber > 0 && !pageNumbers.ContainsKey(objectNumber)) pageNumbers.Add(objectNumber, pageNumbers.Count + 1);
            return;
        }
        if (!dictionary.Items.TryGetValue("Kids", out PdfObject? kidsObject) ||
            PdfObjectLookup.Resolve(objects, kidsObject) is not PdfArray kids) return;
        for (int index = 0; index < kids.Items.Count; index++) {
            AddPageTreeNode(kids.Items[index], objects, visited, pageNumbers, depth + 1, limits, ref complete);
        }
    }

    private static void AppendText(StringBuilder builder, char prefix, string value) =>
        builder.Append(prefix).Append(value.Length).Append(':').Append(value);

    private sealed class PageNumberLookup {
        internal PageNumberLookup(IReadOnlyDictionary<int, int> value, bool isComplete) { Value = value; IsComplete = isComplete; }
        internal IReadOnlyDictionary<int, int> Value { get; }
        internal bool IsComplete { get; }
    }

    private sealed class PageNumberLookupCache {
        private readonly Dictionary<int, PdfIndirectObject> _objects;
        private readonly Dictionary<(int Depth, int Nodes), PageNumberLookup> _values = new();

        internal PageNumberLookupCache(Dictionary<int, PdfIndirectObject> objects) { _objects = objects; }

        internal PageNumberLookup Get(PdfReadLimits limits) {
            var key = (limits.MaxPageTreeDepth, limits.MaxPageTreeNodes);
            lock (_values) {
                if (!_values.TryGetValue(key, out PageNumberLookup? value)) {
                    value = BuildPageNumberLookup(_objects, limits);
                    _values.Add(key, value);
                }
                return value;
            }
        }
    }
}
