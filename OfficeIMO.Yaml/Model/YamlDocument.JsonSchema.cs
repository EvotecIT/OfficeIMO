using OfficeIMO.Json;

namespace OfficeIMO.Yaml;

public sealed partial class YamlDocument {
    /// <summary>
    /// Converts the first document's root to the JSON data model, with source positions; an empty stream converts to
    /// null. See <see cref="YamlNode.ToJsonData"/>.
    /// </summary>
    /// <exception cref="InvalidOperationException">The value nests deeper than 512 levels.</exception>
    public JsonData ToJsonData() => YamlJsonData.Convert(Root);

    /// <summary>
    /// Validates the first document's root against a JSON Schema. Findings carry the JSON Pointer, line and column of
    /// each failing value. For a stream of several documents, such as Kubernetes manifests, validate each of
    /// <see cref="Documents"/> with <see cref="YamlDocumentPart.Validate"/>.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="schema"/> is null.</exception>
    /// <exception cref="InvalidOperationException">The value nests deeper than 512 levels.</exception>
    public JsonSchemaValidationResult Validate(JsonSchema schema) {
        if (schema is null) throw new ArgumentNullException(nameof(schema));
        return schema.Validate(ToJsonData());
    }
}
