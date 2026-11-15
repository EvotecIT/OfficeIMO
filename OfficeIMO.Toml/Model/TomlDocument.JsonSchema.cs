using OfficeIMO.Json;

namespace OfficeIMO.Toml;

public sealed partial class TomlDocument {
    /// <summary>Converts the document to the JSON data model, with source positions. See <see cref="TomlNode.ToJsonData"/>.</summary>
    /// <exception cref="InvalidOperationException">The value nests deeper than 512 levels.</exception>
    public JsonData ToJsonData() => TomlJsonData.Convert(_root);

    /// <summary>Validates the document against a JSON Schema. Findings carry the JSON Pointer, line and column of each failing value.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="schema"/> is null.</exception>
    /// <exception cref="InvalidOperationException">The value nests deeper than 512 levels.</exception>
    public JsonSchemaValidationResult Validate(JsonSchema schema) {
        if (schema is null) throw new ArgumentNullException(nameof(schema));
        return schema.Validate(ToJsonData());
    }
}
