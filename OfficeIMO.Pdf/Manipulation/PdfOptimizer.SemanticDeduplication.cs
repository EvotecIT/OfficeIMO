using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using OfficeIMO.Pdf.Filters;

namespace OfficeIMO.Pdf;

internal static partial class PdfOptimizer {
    private static void DeduplicateImages(Dictionary<int, PdfIndirectObject> objects, PdfOptimizationOptions options, List<PdfOptimizationAction> actions, List<PdfOptimizationSkippedAction> skippedActions) {
        var groups = new Dictionary<string, List<int>>(StringComparer.Ordinal);
        long totalDecodedBytes = 0;
        foreach (KeyValuePair<int, PdfIndirectObject> entry in objects.OrderBy(static item => item.Key)) {
            if (entry.Value.Value is not PdfStream stream || !string.Equals(ReadName(stream.Dictionary, "Subtype"), "Image", StringComparison.Ordinal)) continue;
            long remainingDecodedBytes = options.MaximumTotalDecodedImageBytes - totalDecodedBytes;
            if (remainingDecodedBytes <= 0) {
                skippedActions.Add(new PdfOptimizationSkippedAction("DeduplicateImage", entry.Key, stream.Data.LongLength, "AggregateDecodeLimit", "Stopped semantic image deduplication after the aggregate decoded-image budget was exhausted."));
                break;
            }
            int maximumDecodedBytes = (int)Math.Min(options.MaximumDecodedImageBytes, Math.Min(int.MaxValue, remainingDecodedBytes));
            if (!StreamDecoder.TryDecode(stream.Dictionary, stream.Data, maximumDecodedBytes, out byte[] decoded, objects)) {
                string reason = maximumDecodedBytes < options.MaximumDecodedImageBytes
                    ? "AggregateDecodeLimit"
                    : "UnsupportedImageFilter";
                string description = reason == "AggregateDecodeLimit"
                    ? "Stopped semantic image deduplication before decoding an image beyond the remaining aggregate budget."
                    : "Skipped semantic image deduplication because the image samples could not be decoded within the configured limit.";
                skippedActions.Add(new PdfOptimizationSkippedAction("DeduplicateImage", entry.Key, stream.Data.LongLength, reason, description));
                if (reason == "AggregateDecodeLimit") break;
                continue;
            }
            if (decoded.LongLength > remainingDecodedBytes) {
                // Defensive fallback for decoders that cannot enforce their output limit incrementally.
                skippedActions.Add(new PdfOptimizationSkippedAction("DeduplicateImage", entry.Key, stream.Data.LongLength, "AggregateDecodeLimit", "Stopped semantic image deduplication after the aggregate decoded-image budget was exhausted."));
                break;
            }
            totalDecodedBytes += decoded.LongLength;
            byte[] digest = ComputeSha256(decoded);
            string fingerprint = BuildCanonicalDictionary(stream.Dictionary, "Length", "Filter", "DecodeParms") + "|sha256:" + Convert.ToBase64String(digest);
            AddGroup(groups, fingerprint, entry.Key);
        }
        ApplyDuplicateGroups(objects, groups, "DeduplicateImage", "Reused a losslessly equivalent decoded image XObject.", actions);
    }

    private static byte[] ComputeSha256(byte[] value) {
#if NET6_0_OR_GREATER
        return SHA256.HashData(value);
#else
        using (SHA256 sha256 = SHA256.Create()) {
            return sha256.ComputeHash(value);
        }
#endif
    }

    private static void DeduplicateTypedDictionaries(Dictionary<int, PdfIndirectObject> objects, string typeName, string actionKind, List<PdfOptimizationAction> actions) {
        var groups = new Dictionary<string, List<int>>(StringComparer.Ordinal);
        foreach (KeyValuePair<int, PdfIndirectObject> entry in objects.OrderBy(static item => item.Key)) {
            if (entry.Value.Value is PdfDictionary dictionary && string.Equals(ReadName(dictionary, "Type"), typeName, StringComparison.Ordinal)) AddGroup(groups, BuildCanonicalObject(dictionary), entry.Key);
        }
        ApplyDuplicateGroups(objects, groups, actionKind, "Reused an identical " + typeName.ToLowerInvariant() + " dictionary.", actions);
    }

    private static void DeduplicateResourceDictionaries(Dictionary<int, PdfIndirectObject> objects, List<PdfOptimizationAction> actions) {
        var resourceObjectNumbers = new HashSet<int>();
        foreach (PdfIndirectObject indirect in objects.Values) CollectResourceReferences(indirect.Value, resourceObjectNumbers);
        var groups = new Dictionary<string, List<int>>(StringComparer.Ordinal);
        foreach (int objectNumber in resourceObjectNumbers.OrderBy(static value => value)) {
            if (objects.TryGetValue(objectNumber, out PdfIndirectObject? indirect) && indirect.Value is PdfDictionary dictionary) AddGroup(groups, BuildCanonicalObject(dictionary), objectNumber);
        }
        ApplyDuplicateGroups(objects, groups, "DeduplicateResource", "Reused an identical indirect resource dictionary.", actions);
    }

    private static void CollectResourceReferences(PdfObject value, HashSet<int> result) {
        if (value is PdfArray array) { for (int i = 0; i < array.Items.Count; i++) CollectResourceReferences(array.Items[i], result); return; }
        PdfDictionary? dictionary = value is PdfDictionary direct ? direct : value is PdfStream stream ? stream.Dictionary : null;
        if (dictionary is null) return;
        foreach (KeyValuePair<string, PdfObject> entry in dictionary.Items) {
            if (string.Equals(entry.Key, "Resources", StringComparison.Ordinal) && entry.Value is PdfReference reference) result.Add(reference.ObjectNumber);
            if (entry.Value is not PdfReference) CollectResourceReferences(entry.Value, result);
        }
    }

    private static void AddGroup(Dictionary<string, List<int>> groups, string fingerprint, int objectNumber) {
        if (!groups.TryGetValue(fingerprint, out List<int>? group)) { group = new List<int>(); groups.Add(fingerprint, group); }
        group.Add(objectNumber);
    }

    private static void ApplyDuplicateGroups(Dictionary<int, PdfIndirectObject> objects, Dictionary<string, List<int>> groups, string actionKind, string description, List<PdfOptimizationAction> actions) {
        var replacements = new Dictionary<int, PdfReference>();
        foreach (List<int> group in groups.Values) {
            if (group.Count < 2) continue;
            PdfIndirectObject keeper = objects[group[0]];
            for (int i = 1; i < group.Count; i++) replacements[group[i]] = new PdfReference(keeper.ObjectNumber, keeper.Generation);
        }
        if (replacements.Count == 0) return;
        foreach (int objectNumber in objects.Keys.OrderBy(static value => value).ToArray()) {
            PdfIndirectObject indirect = objects[objectNumber];
            PdfObject rewritten = ReplaceReferences(indirect.Value, replacements);
            if (!ReferenceEquals(rewritten, indirect.Value)) objects[objectNumber] = new PdfIndirectObject(indirect.ObjectNumber, indirect.Generation, rewritten);
        }
        foreach (KeyValuePair<int, PdfReference> replacement in replacements.OrderBy(static item => item.Key)) {
            if (!objects.TryGetValue(replacement.Key, out PdfIndirectObject? duplicate)) continue;
            long originalLength = EstimateObjectLength(duplicate); objects.Remove(replacement.Key);
            actions.Add(new PdfOptimizationAction(actionKind, replacement.Key, originalLength, 0, description + " Keeper object: " + replacement.Value.ObjectNumber.ToString(CultureInfo.InvariantCulture) + "."));
        }
    }

    private static string BuildCanonicalDictionary(PdfDictionary dictionary, params string[] excludedKeys) {
        var excluded = new HashSet<string>(excludedKeys, StringComparer.Ordinal);
        var builder = new StringBuilder(); AppendCanonicalDictionary(builder, dictionary, excluded); return builder.ToString();
    }

    private static string BuildCanonicalObject(PdfObject value) { var builder = new StringBuilder(); AppendCanonical(builder, value, null); return builder.ToString(); }

    private static void AppendCanonical(StringBuilder builder, PdfObject value, HashSet<string>? excludedKeys) {
        switch (value) {
            case PdfNumber number: builder.Append(number.Value.ToString("R", CultureInfo.InvariantCulture)); break;
            case PdfBoolean boolean: builder.Append(boolean.Value ? "true" : "false"); break;
            case PdfName name: builder.Append('/').Append(name.Name); break;
            case PdfStringObj text: builder.Append('<').Append(Convert.ToBase64String(text.RawBytes)).Append('>'); break;
            case PdfReference reference: builder.Append(reference.ObjectNumber.ToString(CultureInfo.InvariantCulture)).Append(':').Append(reference.Generation.ToString(CultureInfo.InvariantCulture)).Append('R'); break;
            case PdfArray array:
                builder.Append('['); for (int i = 0; i < array.Items.Count; i++) { AppendCanonical(builder, array.Items[i], excludedKeys); builder.Append(';'); } builder.Append(']'); break;
            case PdfDictionary dictionary: AppendCanonicalDictionary(builder, dictionary, excludedKeys); break;
            case PdfStream stream: AppendCanonicalDictionary(builder, stream.Dictionary, excludedKeys); builder.Append('|').Append(Convert.ToBase64String(stream.Data)); break;
            case PdfNull: builder.Append("null"); break;
        }
    }

    private static void AppendCanonicalDictionary(StringBuilder builder, PdfDictionary dictionary, HashSet<string>? excludedKeys) {
        builder.Append("<<");
        foreach (KeyValuePair<string, PdfObject> entry in dictionary.Items.OrderBy(static item => item.Key, StringComparer.Ordinal)) {
            if (excludedKeys?.Contains(entry.Key) == true) continue;
            builder.Append('/').Append(entry.Key).Append('='); AppendCanonical(builder, entry.Value, excludedKeys); builder.Append(';');
        }
        builder.Append(">>");
    }

    private static string? ReadName(PdfDictionary dictionary, string key) => dictionary.Items.TryGetValue(key, out PdfObject? value) && value is PdfName name ? name.Name : null;
}
