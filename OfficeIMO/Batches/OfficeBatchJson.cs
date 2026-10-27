using System.Text.Json;

namespace OfficeIMO;

internal static class OfficeBatchJson {
    internal static void CheckEnvironmentName(string name) {
        if (name.Length is < 1 or > 128 || !((name[0] >= 'a' && name[0] <= 'z' || name[0] >= 'A' && name[0] <= 'Z') || name[0] == '_') ||
            name.Any(character => !((character >= 'a' && character <= 'z' || character >= 'A' && character <= 'Z' || character >= '0' && character <= '9') || character == '_')))
            throw new InvalidDataException("Secret references must be environment-variable names.");
    }
    internal static void RejectDuplicates(JsonElement element) {
        if (element.ValueKind == JsonValueKind.Object) {
            var names = new HashSet<string>(StringComparer.Ordinal);
            foreach (JsonProperty property in element.EnumerateObject()) {
                if (!names.Add(property.Name)) throw new JsonException("Duplicate batch property: " + property.Name);
                RejectDuplicates(property.Value);
            }
        } else if (element.ValueKind == JsonValueKind.Array)
            foreach (JsonElement item in element.EnumerateArray()) RejectDuplicates(item);
    }
}
