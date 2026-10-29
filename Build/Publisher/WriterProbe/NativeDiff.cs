using System.Text.Json;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Compares existing bounded inventory views, separating moved records from changed values.</summary>
internal static class NativeDiff {
    internal static object Compare(string leftPath, string rightPath) {
        JsonElement left = JsonSerializer.SerializeToElement(NativeInventory.Read(leftPath), Program.JsonOptions);
        JsonElement right = JsonSerializer.SerializeToElement(NativeInventory.Read(rightPath), Program.JsonOptions);
        var leftStreams = Index(left.GetProperty("streams"), element => element.GetProperty("path").GetString()!);
        var rightStreams = Index(right.GetProperty("streams"), element => element.GetProperty("path").GetString()!);
        var leftChunks = ContentsChunks(left);
        var rightChunks = ContentsChunks(right);
        return new {
            schema = "officeimo.publisher.writer-probe.diff.v1",
            left = new { path = leftPath, sha256 = left.GetProperty("sha256") },
            right = new { path = rightPath, sha256 = right.GetProperty("sha256") },
            rootClassId = new { left = left.GetProperty("rootClassId"), right = right.GetProperty("rootClassId") },
            streams = Changes(leftStreams, rightStreams, "sha256"),
            contentsStatus = new { left = left.GetProperty("contents").GetProperty("status"), right = right.GetProperty("contents").GetProperty("status") },
            contentsHeader = new { left = Scalar(left.GetProperty("contents"), "header"), right = Scalar(right.GetProperty("contents"), "header") },
            contentsChunks = Changes(leftChunks, rightChunks, "sha256"),
            movedContentsChunks = leftChunks.Keys.Intersect(rightChunks.Keys).Where(key =>
                Scalar(leftChunks[key], "Offset") != Scalar(rightChunks[key], "Offset")).Select(key => new {
                    id = key, leftOffset = leftChunks[key].GetProperty("Offset"), rightOffset = rightChunks[key].GetProperty("Offset")
                }).ToArray(),
            fields = FieldChanges(leftChunks, rightChunks),
            quillStatus = new { left = left.GetProperty("quill").GetProperty("status"), right = right.GetProperty("quill").GetProperty("status") },
            quillChunks = Changes(QuillChunks(left), QuillChunks(right), "sha256"),
            limitation = "Differences describe inspected records and opaque stream identities. Equal inspected values do not establish native semantic equivalence or visual fidelity."
        };
    }

    private static Dictionary<string, JsonElement> Index(JsonElement array, Func<JsonElement, string> key) =>
        array.EnumerateArray().ToDictionary(key, element => element, StringComparer.Ordinal);

    private static Dictionary<string, JsonElement> ContentsChunks(JsonElement inventory) {
        JsonElement contents = inventory.GetProperty("contents");
        return contents.TryGetProperty("chunks", out var chunks)
            ? Index(chunks, element => element.GetProperty("Id").GetUInt32().ToString(System.Globalization.CultureInfo.InvariantCulture))
            : new(StringComparer.Ordinal);
    }

    private static Dictionary<string, JsonElement> QuillChunks(JsonElement inventory) {
        JsonElement quill = inventory.GetProperty("quill");
        return quill.TryGetProperty("chunks", out var chunks)
            ? Index(chunks, element => element.GetProperty("Tag").GetString() + ":" + element.GetProperty("Id").GetInt32())
            : new(StringComparer.Ordinal);
    }

    private static object[] Changes(Dictionary<string, JsonElement> left, Dictionary<string, JsonElement> right, string fingerprint) =>
        left.Keys.Union(right.Keys).OrderBy(key => key, StringComparer.Ordinal).Where(key =>
            !left.ContainsKey(key) || !right.ContainsKey(key) || Scalar(left[key], fingerprint) != Scalar(right[key], fingerprint))
            .Select(key => (object)new {
                key, change = !left.ContainsKey(key) ? "added" : !right.ContainsKey(key) ? "removed" : "changed",
                left = left.TryGetValue(key, out var l) ? Compact(l) : null,
                right = right.TryGetValue(key, out var r) ? Compact(r) : null
            }).ToArray();

    private static Dictionary<string, JsonElement> Compact(JsonElement value) => value.EnumerateObject()
        .Where(property => property.Name is not ("fields" or "prefix"))
        .ToDictionary(property => property.Name, property => property.Value, StringComparer.Ordinal);

    private static string? Scalar(JsonElement element, string name) =>
        element.TryGetProperty(name, out var value) ? value.ToString() : null;

    private static object[] FieldChanges(Dictionary<string, JsonElement> leftChunks, Dictionary<string, JsonElement> rightChunks) {
        var left = Flatten(leftChunks);
        var right = Flatten(rightChunks);
        return left.Keys.Union(right.Keys).OrderBy(key => key, StringComparer.Ordinal).Where(key =>
            !left.ContainsKey(key) || !right.ContainsKey(key) || left[key] != right[key]).Select(key => (object)new {
                path = key, left = left.GetValueOrDefault(key), right = right.GetValueOrDefault(key)
            }).ToArray();
    }

    private static Dictionary<string, string> Flatten(Dictionary<string, JsonElement> chunks) {
        var result = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var entry in chunks) {
            var fields = entry.Value.GetProperty("fields");
            if (fields.TryGetProperty("blocks", out var blocks)) Visit(blocks, "chunk/" + entry.Key);
        }
        return result;

        void Visit(JsonElement blocks, string path) {
            var occurrences = new Dictionary<string, int>(StringComparer.Ordinal);
            foreach (var block in blocks.EnumerateArray()) {
                if (!block.TryGetProperty("Id", out var id)) continue;
                string key = id.GetByte().ToString("X2") + ":" + block.GetProperty("Type").GetByte().ToString("X2");
                int occurrence = occurrences.GetValueOrDefault(key);
                occurrences[key] = occurrence + 1;
                string fieldPath = path + "/" + key + "[" + occurrence + "]";
                result[fieldPath] = string.Join("|", Scalar(block, "Value"), Scalar(block, "text"), Scalar(block, "rawValue"));
                if (block.TryGetProperty("children", out var children) && children.ValueKind == JsonValueKind.Object
                    && children.TryGetProperty("blocks", out var nested)) Visit(nested, fieldPath);
            }
        }
    }
}
