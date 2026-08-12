using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace OfficeIMO.Pdf;

internal static class PdfActionPayloadFingerprint {
    private const int MaximumDepth = 32;
    private const int MaximumNodes = 4096;

    internal static string Create(
        PdfDictionary action,
        Dictionary<int, PdfIndirectObject> objects) {
        var builder = new StringBuilder();
        var activeReferences = new HashSet<(int ObjectNumber, int Generation)>();
        int nodes = 0;
        AppendDictionary(builder, action, objects, activeReferences, depth: 0, ref nodes, isActionRoot: true);
        return builder.ToString();
    }

    private static void AppendObject(
        StringBuilder builder,
        PdfObject? value,
        Dictionary<int, PdfIndirectObject> objects,
        HashSet<(int ObjectNumber, int Generation)> activeReferences,
        int depth,
        ref int nodes) {
        nodes++;
        if (depth > MaximumDepth || nodes > MaximumNodes) {
            builder.Append("!limit");
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
                AppendReference(builder, reference, objects, activeReferences, depth, ref nodes);
                return;
            case PdfArray array:
                builder.Append('[');
                for (int i = 0; i < array.Items.Count; i++) {
                    AppendObject(builder, array.Items[i], objects, activeReferences, depth + 1, ref nodes);
                    builder.Append(';');
                }
                builder.Append(']');
                return;
            case PdfDictionary dictionary:
                AppendDictionary(builder, dictionary, objects, activeReferences, depth + 1, ref nodes, isActionRoot: false);
                return;
            case PdfStream stream:
                builder.Append("stream:");
                AppendDictionary(builder, stream.Dictionary, objects, activeReferences, depth + 1, ref nodes, isActionRoot: false);
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
        HashSet<(int ObjectNumber, int Generation)> activeReferences,
        int depth,
        ref int nodes) {
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
                builder.Append("page:").Append(reference.ObjectNumber).Append(':').Append(reference.Generation);
                return;
            }
            AppendObject(builder, indirect.Value, objects, activeReferences, depth + 1, ref nodes);
        } finally {
            activeReferences.Remove(key);
        }
    }

    private static void AppendDictionary(
        StringBuilder builder,
        PdfDictionary dictionary,
        Dictionary<int, PdfIndirectObject> objects,
        HashSet<(int ObjectNumber, int Generation)> activeReferences,
        int depth,
        ref int nodes,
        bool isActionRoot) {
        builder.Append('{');
        foreach (string key in dictionary.Items.Keys.OrderBy(static key => key, StringComparer.Ordinal)) {
            if (isActionRoot &&
                (string.Equals(key, "S", StringComparison.Ordinal) ||
                 string.Equals(key, "Next", StringComparison.Ordinal))) continue;
            AppendText(builder, 'K', key);
            AppendObject(builder, dictionary.Items[key], objects, activeReferences, depth + 1, ref nodes);
            builder.Append(';');
        }
        builder.Append('}');
    }

    private static void AppendText(StringBuilder builder, char prefix, string value) =>
        builder.Append(prefix).Append(value.Length).Append(':').Append(value);
}
