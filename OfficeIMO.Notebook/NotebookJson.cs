using System.Text;

namespace OfficeIMO.Notebook;

internal static class NotebookJson {
    internal static JsonElement Property(JsonElement owner, string name) =>
        owner.ValueKind == JsonValueKind.Object && owner.TryGetProperty(name, out JsonElement value) ? value : default;

    internal static string? String(JsonElement owner, string name) {
        JsonElement property = Property(owner, name);
        return property.ValueKind == JsonValueKind.String ? property.GetString() : null;
    }

    internal static int? Integer(JsonElement owner, string name) {
        JsonElement property = Property(owner, name);
        return property.ValueKind == JsonValueKind.Number && property.TryGetInt32(out int value) ? value : null;
    }

    internal static Dictionary<string, JsonElement> ReadMap(JsonElement value, Func<string, string>? mapName = null) {
        var map = new Dictionary<string, JsonElement>(StringComparer.Ordinal);
        if (value.ValueKind == JsonValueKind.Object) {
            foreach (JsonProperty property in value.EnumerateObject()) {
                map[mapName?.Invoke(property.Name) ?? property.Name] = property.Value;
            }
        }
        return map;
    }

    internal static Dictionary<string, JsonElement> CopyMap(IReadOnlyDictionary<string, JsonElement>? source) {
        var map = new Dictionary<string, JsonElement>(StringComparer.Ordinal);
        if (source != null) {
            foreach (KeyValuePair<string, JsonElement> property in source) {
                map.Add(property.Key, CloneValue(property.Value));
            }
        }
        return map;
    }

    internal static JsonElement CloneValue(JsonElement value) {
        if (value.ValueKind == JsonValueKind.Undefined) throw new ArgumentException("A defined JSON value is required.", nameof(value));
        return value.Clone();
    }

    internal static string ReadText(JsonElement value) => ReadText(value, int.MaxValue, out _);

    internal static string ReadText(JsonElement value, int maxChars, out bool truncated) {
        truncated = false;
        if (value.ValueKind == JsonValueKind.String) {
            return LimitText(value.GetString()!, maxChars, out truncated);
        }
        if (value.ValueKind != JsonValueKind.Array) return string.Empty;
        var text = new StringBuilder(Math.Min(maxChars, 512));
        foreach (JsonElement part in value.EnumerateArray()) {
            if (part.ValueKind != JsonValueKind.String) continue;
            string fragment = LimitText(part.GetString()!, maxChars - text.Length, out truncated);
            text.Append(fragment);
            if (truncated) break;
        }
        return text.ToString();
    }

    internal static string LimitText(string text, int maxChars, out bool truncated) {
        truncated = text.Length > maxChars;
        if (!truncated) return text;
        if (maxChars > 0 && char.IsHighSurrogate(text[maxChars - 1])) maxChars--;
        return text.Substring(0, maxChars);
    }

    internal static bool WriteOriginal(Utf8JsonWriter writer, string name, JsonElement value) {
        if (value.ValueKind == JsonValueKind.Undefined) return false;
        writer.WritePropertyName(name);
        value.WriteTo(writer);
        return true;
    }

    internal static void WriteMap(Utf8JsonWriter writer, string name, IReadOnlyDictionary<string, JsonElement> map) {
        writer.WriteStartObject(name);
        foreach (KeyValuePair<string, JsonElement> property in map) {
            writer.WritePropertyName(property.Key);
            property.Value.WriteTo(writer);
        }
        writer.WriteEndObject();
    }

    internal static void CopyProperties(Utf8JsonWriter writer, JsonElement owner, Func<JsonProperty, bool> include) {
        if (owner.ValueKind != JsonValueKind.Object) return;
        foreach (JsonProperty property in owner.EnumerateObject()) {
            if (include(property)) property.WriteTo(writer);
        }
    }
}
