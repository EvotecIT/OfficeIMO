using System.IO;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;

namespace OfficeIMO.Adf;

/// <summary>Validates the vocabulary used by the bundled Atlassian Draft 4 schema without another runtime dependency.</summary>
internal static class AdfPinnedSchema {
    private static readonly Lazy<JsonDocument> Schema = new Lazy<JsonDocument>(Load);

    internal static AdfValidationResult Validate(AdfDocument document, long maximumEvaluations, AdfProcessingOptions options) {
        AdfValidationIssue? graphIssue = AdfGraphSafety.Inspect(document, options);
        if (graphIssue != null) return new AdfValidationResult(new[] { graphIssue });
        CancellationToken cancellationToken = options.CancellationToken;
        int maximumJsonDepth = options.MaxDepth * 3 + 8;
        string json;
        try { json = AdfJsonSerializer.Serialize(document, false, options); }
        catch (InvalidDataException exception) { return new AdfValidationResult(new[] { new AdfValidationIssue("ADF_OUTPUT_LIMIT_EXCEEDED", "$", exception.Message, AdfValidationSeverity.Error) }); }
        using JsonDocument value = JsonDocument.Parse(json, new JsonDocumentOptions { MaxDepth = maximumJsonDepth });
        cancellationToken.ThrowIfCancellationRequested();
        var issues = new List<AdfValidationIssue>();
        try { new Walker(maximumEvaluations, maximumJsonDepth, cancellationToken).Check(Schema.Value.RootElement, value.RootElement, "$", issues, 0); }
        catch (EvaluationLimitException exception) { issues.Add(new AdfValidationIssue("ADF_SCHEMA_LIMIT", exception.Path, "Full-schema validation exceeds MaximumSchemaEvaluations.", AdfValidationSeverity.Error)); }
        return new AdfValidationResult(issues);
    }

    private static JsonDocument Load() {
        using Stream stream = typeof(AdfPinnedSchema).Assembly.GetManifestResourceStream("OfficeIMO.Adf.Schema.adf-schema.json")
            ?? throw new InvalidOperationException("The bundled ADF schema resource is missing.");
        return JsonDocument.Parse(stream);
    }

    private sealed class Walker {
        private readonly long _maximumEvaluations;
        private readonly int _maximumJsonDepth;
        private readonly CancellationToken _cancellationToken;
        private long _evaluations;

        internal Walker(long maximumEvaluations, int maximumJsonDepth, CancellationToken cancellationToken) { _maximumEvaluations = maximumEvaluations; _maximumJsonDepth = maximumJsonDepth; _cancellationToken = cancellationToken; }

        internal void Check(JsonElement rule, JsonElement value, string path, List<AdfValidationIssue> issues, int depth) {
            _cancellationToken.ThrowIfCancellationRequested();
            if (_evaluations++ >= _maximumEvaluations) throw new EvaluationLimitException(path);
            if (issues.Count >= 1000) return;
            if (depth > _maximumJsonDepth) { Error(issues, path, "JSON data exceeds the validation depth limit."); return; }
            if (rule.TryGetProperty("$ref", out JsonElement reference)) {
                Check(Resolve(reference.GetString()!), value, path, issues, depth);
                return;
            }
            if (rule.TryGetProperty("allOf", out JsonElement all))
                foreach (JsonElement child in all.EnumerateArray()) Check(child, value, path, issues, depth);
            if (rule.TryGetProperty("anyOf", out JsonElement alternatives)) {
                List<AdfValidationIssue>? best = null;
                bool matched = false;
                foreach (JsonElement child in alternatives.EnumerateArray()) {
                    // A different explicit node discriminator cannot match. Do not descend into
                    // that branch's content: recursive unions otherwise multiply validation work.
                    if (ExcludesNodeType(child, value)) continue;
                    var candidate = new List<AdfValidationIssue>();
                    Check(child, value, path, candidate, depth);
                    if (candidate.Count == 0) { matched = true; break; }
                    if (best == null || candidate.Count < best.Count) best = candidate;
                }
                if (!matched) {
                    if (best != null) issues.AddRange(best.Take(Math.Max(0, 1000 - issues.Count)));
                    else Error(issues, path + ".type", "The node type is not allowed at this position by the pinned schema.");
                }
            }
            if (rule.TryGetProperty("type", out JsonElement type) && !HasType(value, type.GetString()!)) {
                Error(issues, path, "Expected a JSON " + type.GetString() + " value.");
                return;
            }
            if (rule.TryGetProperty("enum", out JsonElement values) && !values.EnumerateArray().Any(item => Equal(item, value)))
                Error(issues, path, "The value is outside the schema enumeration.");
            if (value.ValueKind == JsonValueKind.Object) CheckObject(rule, value, path, issues, depth);
            if (value.ValueKind == JsonValueKind.Array) CheckArray(rule, value, path, issues, depth);
            if (value.ValueKind == JsonValueKind.Number && value.TryGetDouble(out double number)) {
                if (rule.TryGetProperty("minimum", out JsonElement minimum) && number < minimum.GetDouble()) Error(issues, path, "The number is below the schema minimum.");
                if (rule.TryGetProperty("maximum", out JsonElement maximum) && number > maximum.GetDouble()) Error(issues, path, "The number exceeds the schema maximum.");
            }
            if (value.ValueKind == JsonValueKind.String) {
                string text = value.GetString()!;
                if (rule.TryGetProperty("minLength", out JsonElement minimum) && CodePointLength(text) < minimum.GetInt32()) Error(issues, path, "The string is shorter than the schema minimum.");
                if (rule.TryGetProperty("pattern", out JsonElement pattern) && !Regex.IsMatch(text, pattern.GetString()!, RegexOptions.CultureInvariant, TimeSpan.FromMilliseconds(100)))
                    Error(issues, path, "The string does not match the schema pattern.");
            }
        }

        private void CheckObject(JsonElement rule, JsonElement value, string path, List<AdfValidationIssue> issues, int depth) {
            if (rule.TryGetProperty("required", out JsonElement required)) {
                foreach (JsonElement name in required.EnumerateArray())
                    if (!value.TryGetProperty(name.GetString()!, out _)) Error(issues, path + "." + name.GetString(), "A required property is missing.");
            }
            bool hasProperties = rule.TryGetProperty("properties", out JsonElement properties);
            bool allowExtra = !rule.TryGetProperty("additionalProperties", out JsonElement additional) || additional.ValueKind != JsonValueKind.False;
            foreach (JsonProperty property in value.EnumerateObject()) {
                _cancellationToken.ThrowIfCancellationRequested();
                if (hasProperties && properties.TryGetProperty(property.Name, out JsonElement child))
                    Check(child, property.Value, path + "." + property.Name, issues, depth + 1);
                else if (!allowExtra) Error(issues, path + "." + property.Name, "The property is not allowed by the pinned schema.");
            }
        }

        private void CheckArray(JsonElement rule, JsonElement value, string path, List<AdfValidationIssue> issues, int depth) {
            int count = value.GetArrayLength();
            if (rule.TryGetProperty("minItems", out JsonElement minimum) && count < minimum.GetInt32()) Error(issues, path, "The array has fewer items than the schema requires.");
            if (rule.TryGetProperty("maxItems", out JsonElement maximum) && count > maximum.GetInt32()) Error(issues, path, "The array has more items than the schema allows.");
            if (!rule.TryGetProperty("items", out JsonElement items)) return;
            int index = 0;
            foreach (JsonElement item in value.EnumerateArray()) {
                _cancellationToken.ThrowIfCancellationRequested();
                if (items.ValueKind != JsonValueKind.Array) Check(items, item, path + "[" + index + "]", issues, depth + 1);
                else if (index < items.GetArrayLength()) Check(items[index], item, path + "[" + index + "]", issues, depth + 1);
                index++;
            }
        }
    }

    private static bool ExcludesNodeType(JsonElement rule, JsonElement value) {
        if (value.ValueKind != JsonValueKind.Object || !value.TryGetProperty("type", out JsonElement nodeType)) return false;
        if (rule.TryGetProperty("$ref", out JsonElement reference)) return ExcludesNodeType(Resolve(reference.GetString()!), value);
        if (rule.TryGetProperty("allOf", out JsonElement all)) return all.EnumerateArray().Any(child => ExcludesNodeType(child, value));
        return rule.TryGetProperty("properties", out JsonElement properties) && properties.TryGetProperty("type", out JsonElement type) &&
            type.TryGetProperty("enum", out JsonElement values) && !values.EnumerateArray().Any(item => Equal(item, nodeType));
    }

    private static JsonElement Resolve(string reference) {
        const string prefix = "#/definitions/";
        if (!reference.StartsWith(prefix, StringComparison.Ordinal)) throw new InvalidOperationException("The bundled ADF schema contains an unsupported reference.");
        return Schema.Value.RootElement.GetProperty("definitions").GetProperty(reference.Substring(prefix.Length));
    }

    private static bool HasType(JsonElement value, string type) => type switch {
        "object" => value.ValueKind == JsonValueKind.Object,
        "array" => value.ValueKind == JsonValueKind.Array,
        "string" => value.ValueKind == JsonValueKind.String,
        "number" => value.ValueKind == JsonValueKind.Number,
        "boolean" => value.ValueKind == JsonValueKind.True || value.ValueKind == JsonValueKind.False,
        "null" => value.ValueKind == JsonValueKind.Null,
        _ => throw new InvalidOperationException("The bundled ADF schema contains an unsupported type.")
    };

    private static bool Equal(JsonElement first, JsonElement second) => first.ValueKind == second.ValueKind &&
        (first.ValueKind == JsonValueKind.String ? first.GetString() == second.GetString() :
        first.ValueKind == JsonValueKind.Number && first.TryGetDecimal(out decimal a) && second.TryGetDecimal(out decimal b) ? a == b : first.GetRawText() == second.GetRawText());
    private static int CodePointLength(string value) { int length = 0; for (int index = 0; index < value.Length; index++, length++) if (char.IsHighSurrogate(value[index]) && index + 1 < value.Length && char.IsLowSurrogate(value[index + 1])) index++; return length; }
    private static void Error(List<AdfValidationIssue> issues, string path, string message) {
        if (issues.Count < 1000) issues.Add(new AdfValidationIssue("ADF_SCHEMA", path, message, AdfValidationSeverity.Error));
    }
    private sealed class EvaluationLimitException : Exception {
        internal EvaluationLimitException(string path) { Path = path; }
        internal string Path { get; }
    }
}
