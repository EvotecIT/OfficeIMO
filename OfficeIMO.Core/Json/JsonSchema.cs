using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeIMO.Json;

/// <summary>
/// A JSON Schema (draft 2020-12 or draft 7) that validates <see cref="JsonData"/> values, such as the contents of a
/// YAML, TOML or JSON configuration file. Findings carry the JSON Pointer of the value, the schema path of the failing
/// keyword, and the value's line and column in the source.
/// </summary>
/// <remarks>
/// <para>
/// Supported keywords: <c>type</c>, <c>enum</c>, <c>const</c>; <c>multipleOf</c>, <c>minimum</c>, <c>maximum</c>,
/// <c>exclusiveMinimum</c>, <c>exclusiveMaximum</c>; <c>minLength</c>, <c>maxLength</c>, <c>pattern</c>,
/// <c>format</c>; <c>prefixItems</c>, <c>items</c>, <c>contains</c>, <c>minContains</c>, <c>maxContains</c>,
/// <c>minItems</c>, <c>maxItems</c>, <c>uniqueItems</c>, <c>unevaluatedItems</c>; <c>properties</c>,
/// <c>patternProperties</c>, <c>additionalProperties</c>, <c>propertyNames</c>, <c>required</c>,
/// <c>minProperties</c>, <c>maxProperties</c>, <c>dependentRequired</c>, <c>dependentSchemas</c>,
/// <c>unevaluatedProperties</c>; <c>allOf</c>, <c>anyOf</c>, <c>oneOf</c>, <c>not</c>, <c>if</c>/<c>then</c>/<c>else</c>;
/// <c>$ref</c>, <c>$dynamicRef</c>, <c>$defs</c>, <c>$id</c>, <c>$anchor</c> and <c>$dynamicAnchor</c>. Draft 7 schemas
/// also read <c>definitions</c>, array <c>items</c> with <c>additionalItems</c>, and <c>dependencies</c>. Other
/// keywords, including <c>$vocabulary</c> and the content keywords, are annotations only.
/// </para>
/// <para>
/// References resolve within the loaded document; anything else goes to <see cref="JsonSchemaOptions.ReferenceResolver"/>,
/// and nothing is fetched otherwise. Patterns use .NET regular expressions with a timeout. A loaded schema is safe to
/// use from several threads.
/// </para>
/// </remarks>
public sealed partial class JsonSchema {
    private readonly object _sync = new object();
    private readonly JsonSchemaOptions _options;
    private readonly Dictionary<string, JsonSchemaResource?> _resources = new Dictionary<string, JsonSchemaResource?>(StringComparer.Ordinal);
    private readonly Dictionary<JsonData, JsonSchemaResource> _resourceOf = new Dictionary<JsonData, JsonSchemaResource>(JsonDataReferenceComparer.Instance);
    private readonly Dictionary<JsonData, JsonSchemaNode> _compiled = new Dictionary<JsonData, JsonSchemaNode>(JsonDataReferenceComparer.Instance);
    private readonly JsonSchemaNode _root;

    private JsonSchema(JsonData schema, JsonSchemaOptions options) {
        _options = options;
        Root = schema;
        Draft = options.Draft ?? DetectDraft(schema);
        var resource = new JsonSchemaResource(JsonSchemaUri.ResourcePart(options.BaseUri ?? string.Empty), schema);
        _resources[resource.Uri] = resource;
        Discover(schema, resource, 0);
        _root = Compile(schema, resource);
    }

    /// <summary>Gets the dialect the schema is read with.</summary>
    public JsonSchemaDraft Draft { get; }

    /// <summary>Gets the schema as loaded.</summary>
    public JsonData Root { get; }

    /// <summary>Parses a schema from JSON text.</summary>
    /// <exception cref="FormatException">The text is not JSON.</exception>
    /// <exception cref="JsonSchemaException">The schema is not valid.</exception>
    public static JsonSchema Parse(string json, JsonSchemaOptions? options = null) => Load(JsonData.Parse(json), options);

    /// <summary>Loads a schema from a JSON value, such as one converted from a YAML document.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="schema"/> is null.</exception>
    /// <exception cref="JsonSchemaException">The schema is not an object or boolean, or a keyword has a value of the wrong type.</exception>
    public static JsonSchema Load(JsonData schema, JsonSchemaOptions? options = null) {
        if (schema is null) throw new ArgumentNullException(nameof(schema));
        if (schema.Kind is not (JsonDataKind.Object or JsonDataKind.Boolean)) throw new JsonSchemaException("A schema must be an object or a boolean, not " + JsonData.Describe(schema.Kind) + ".");
        return new JsonSchema(schema, options?.Clone() ?? new JsonSchemaOptions());
    }

    /// <summary>Validates a value, reporting every finding up to <see cref="JsonSchemaOptions.MaxErrors"/>.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="instance"/> is null.</exception>
    /// <exception cref="JsonSchemaException">A referenced part of the schema is not valid.</exception>
    public JsonSchemaValidationResult Validate(JsonData instance) {
        if (instance is null) throw new ArgumentNullException(nameof(instance));
        var context = new EvaluationContext(this, report: true);
        Evaluate(_root, instance, null, null, true, null, null, 0, context);
        return context.ToResult();
    }

    /// <summary>Returns whether a value matches, stopping at the first finding.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="instance"/> is null.</exception>
    /// <exception cref="JsonSchemaException">A referenced part of the schema is not valid.</exception>
    public bool IsValid(JsonData instance) {
        if (instance is null) throw new ArgumentNullException(nameof(instance));
        var context = new EvaluationContext(this, report: false);
        return Evaluate(_root, instance, null, null, false, null, null, 0, context) && !context.Aborted;
    }

    private static JsonSchemaDraft DetectDraft(JsonData schema) {
        if (schema.Kind != JsonDataKind.Object || !schema.TryGetProperty("$schema", out JsonData uri) || uri.Kind != JsonDataKind.String) return JsonSchemaDraft.Draft202012;
        string text = uri.GetString();
        return text.IndexOf("draft-07", StringComparison.Ordinal) >= 0 || text.IndexOf("draft-06", StringComparison.Ordinal) >= 0
            ? JsonSchemaDraft.Draft7
            : JsonSchemaDraft.Draft202012;
    }

    private bool IsDraft7 => Draft == JsonSchemaDraft.Draft7;

    #region Discovery

    /// <summary>Registers embedded resources and anchors, walking only keywords whose values are schemas.</summary>
    private void Discover(JsonData data, JsonSchemaResource resource, int depth) {
        if (data.Kind != JsonDataKind.Object || depth > JsonDataParser.MaxDepth) return;
        if (_resourceOf.ContainsKey(data)) return;
        bool hasRef = data.TryGetProperty("$ref", out _);
        if (data.TryGetProperty("$id", out JsonData id) && id.Kind == JsonDataKind.String && !(IsDraft7 && hasRef)) {
            string text = id.GetString();
            string resolved = JsonSchemaUri.Resolve(resource.Uri, text);
            string resourceUri = JsonSchemaUri.ResourcePart(resolved);
            bool sameResource = text.Length > 0 && text[0] == '#';
            if (IsDraft7 && sameResource) {
                string anchor = JsonSchemaUri.Fragment(resolved);
                if (anchor.Length > 0 && anchor[0] != '/') resource.Anchors[anchor] = data;
            } else if (!sameResource && !JsonSchemaUri.HasFragment(resolved)) {
                JsonSchemaResource named = Register(resourceUri, data);
                // A document whose $id differs from the URI it was loaded under stays reachable under both.
                if (ReferenceEquals(data, resource.Root) && !ReferenceEquals(named, resource)) _resources[resource.Uri] = named;
                resource = named;
            }
        }
        if (!IsDraft7) {
            if (data.TryGetProperty("$anchor", out JsonData anchorName) && anchorName.Kind == JsonDataKind.String) resource.Anchors[anchorName.GetString()] = data;
            if (data.TryGetProperty("$dynamicAnchor", out JsonData dynamicName) && dynamicName.Kind == JsonDataKind.String) {
                resource.Anchors[dynamicName.GetString()] = data;
                (resource.DynamicAnchors ??= new Dictionary<string, JsonData>(StringComparer.Ordinal))[dynamicName.GetString()] = data;
            }
        }
        _resourceOf[data] = resource;
        foreach (KeyValuePair<string, JsonData> property in data.Properties) {
            JsonData value = property.Value;
            switch (property.Key) {
                case "additionalProperties":
                case "propertyNames":
                case "contains":
                case "not":
                case "if":
                case "then":
                case "else":
                case "unevaluatedItems":
                case "unevaluatedProperties":
                case "additionalItems":
                case "contentSchema":
                    Discover(value, resource, depth + 1);
                    break;
                case "items":
                case "allOf":
                case "anyOf":
                case "oneOf":
                case "prefixItems":
                    if (value.Kind == JsonDataKind.Array) {
                        foreach (JsonData item in value.Items) Discover(item, resource, depth + 1);
                    } else {
                        Discover(value, resource, depth + 1);
                    }
                    break;
                case "properties":
                case "patternProperties":
                case "$defs":
                case "definitions":
                case "dependentSchemas":
                case "dependencies":
                    foreach (KeyValuePair<string, JsonData> entry in value.Properties) Discover(entry.Value, resource, depth + 1);
                    break;
            }
        }
    }

    private JsonSchemaResource Register(string uri, JsonData root) {
        if (_resources.TryGetValue(uri, out JsonSchemaResource? existing) && existing is not null && ReferenceEquals(existing.Root, root)) return existing;
        var resource = new JsonSchemaResource(uri, root);
        _resources[uri] = resource;
        return resource;
    }

    #endregion

    #region Compilation

    private JsonSchemaNode Compile(JsonData data, JsonSchemaResource fallback) {
        if (_compiled.TryGetValue(data, out JsonSchemaNode? existing)) return existing;
        JsonSchemaResource resource = _resourceOf.TryGetValue(data, out JsonSchemaResource? owner) ? owner : fallback;
        var node = new JsonSchemaNode(data, resource);
        _compiled[data] = node;
        if (data.Kind == JsonDataKind.Boolean) {
            node.Boolean = data.GetBoolean();
            return node;
        }
        if (data.Kind != JsonDataKind.Object) throw Invalid(data, "A schema must be an object or a boolean, not " + JsonData.Describe(data.Kind));
        bool draft7 = IsDraft7;
        if (data.TryGetProperty("$ref", out JsonData reference)) {
            node.Ref = ReadString(reference, "$ref");
            if (draft7) {
                node.RefOnly = true;
                return node;
            }
        }
        foreach (KeyValuePair<string, JsonData> property in data.Properties) {
            string keyword = property.Key;
            JsonData value = property.Value;
            switch (keyword) {
                case "type": node.Types = ReadTypes(value); break;
                case "enum":
                    if (value.Kind != JsonDataKind.Array) throw Invalid(value, "'enum' must be an array");
                    node.Enum = value.Items.ToArray();
                    break;
                case "const": node.Const = value; break;
                case "multipleOf":
                    node.MultipleOf = ReadNumber(value, keyword);
                    if (node.MultipleOf.DoubleValue <= 0) throw Invalid(value, "'multipleOf' must be greater than 0");
                    break;
                case "maximum": node.Maximum = ReadNumber(value, keyword); break;
                case "minimum": node.Minimum = ReadNumber(value, keyword); break;
                case "exclusiveMaximum": node.ExclusiveMaximum = ReadNumber(value, keyword); break;
                case "exclusiveMinimum": node.ExclusiveMinimum = ReadNumber(value, keyword); break;
                case "maxLength": node.MaxLength = ReadCount(value, keyword); break;
                case "minLength": node.MinLength = ReadCount(value, keyword); break;
                case "pattern":
                    node.PatternText = ReadString(value, keyword);
                    node.Pattern = CreateRegex(node.PatternText, value);
                    break;
                case "format": node.Format = ReadString(value, keyword); break;
                case "maxItems": node.MaxItems = ReadCount(value, keyword); break;
                case "minItems": node.MinItems = ReadCount(value, keyword); break;
                case "uniqueItems": node.UniqueItems = ReadBoolean(value, keyword); break;
                case "maxProperties": node.MaxProperties = ReadCount(value, keyword); break;
                case "minProperties": node.MinProperties = ReadCount(value, keyword); break;
                case "required": node.Required = ReadStrings(value, keyword); break;
                case "items":
                    if (draft7 && value.Kind == JsonDataKind.Array) {
                        node.PrefixItems = CompileArray(value, resource, keyword);
                        node.PrefixItemsKeyword = "items";
                        if (data.TryGetProperty("additionalItems", out JsonData additional)) {
                            node.Items = Compile(additional, resource);
                            node.ItemsKeyword = "additionalItems";
                        }
                    } else if (value.Kind != JsonDataKind.Array) {
                        node.Items = Compile(value, resource);
                    } else {
                        throw Invalid(value, "'items' must be a schema; use 'prefixItems' for a list of schemas");
                    }
                    break;
                case "prefixItems":
                    if (!draft7) node.PrefixItems = CompileArray(value, resource, keyword);
                    break;
                case "contains": node.Contains = Compile(value, resource); break;
                case "maxContains":
                    if (!draft7) node.MaxContains = ReadCount(value, keyword);
                    break;
                case "minContains":
                    if (!draft7) node.MinContains = ReadCount(value, keyword);
                    break;
                case "unevaluatedItems":
                    if (!draft7) node.UnevaluatedItems = Compile(value, resource);
                    break;
                case "properties": node.Properties = CompileMap(value, resource, keyword); break;
                case "patternProperties": {
                    if (value.Kind != JsonDataKind.Object) throw Invalid(value, "'patternProperties' must be an object");
                    var patterns = new JsonSchemaNode.PatternSchema[value.Properties.Count];
                    for (int i = 0; i < patterns.Length; i++) {
                        KeyValuePair<string, JsonData> entry = value.Properties[i];
                        patterns[i] = new JsonSchemaNode.PatternSchema(entry.Key, CreateRegex(entry.Key, entry.Value), Compile(entry.Value, resource));
                    }
                    node.PatternProperties = patterns;
                    break;
                }
                case "additionalProperties": node.AdditionalProperties = Compile(value, resource); break;
                case "propertyNames": node.PropertyNames = Compile(value, resource); break;
                case "unevaluatedProperties":
                    if (!draft7) node.UnevaluatedProperties = Compile(value, resource);
                    break;
                case "dependentRequired":
                    if (!draft7) node.DependentRequired = ReadDependentRequired(value, keyword);
                    break;
                case "dependentSchemas":
                    if (!draft7) node.DependentSchemas = CompileMap(value, resource, keyword);
                    break;
                case "dependencies":
                    // Draft 2020-12 split it into dependentRequired and dependentSchemas; read it there too when they are absent.
                    if (draft7 || (!data.TryGetProperty("dependentRequired", out _) && !data.TryGetProperty("dependentSchemas", out _))) ReadDependencies(node, value, resource);
                    break;
                case "allOf": node.AllOf = CompileArray(value, resource, keyword); break;
                case "anyOf": node.AnyOf = CompileArray(value, resource, keyword); break;
                case "oneOf": node.OneOf = CompileArray(value, resource, keyword); break;
                case "not": node.Not = Compile(value, resource); break;
                case "if": node.If = Compile(value, resource); break;
                case "then": node.Then = Compile(value, resource); break;
                case "else": node.Else = Compile(value, resource); break;
                case "$dynamicRef":
                    if (!draft7) node.DynamicRef = ReadString(value, keyword);
                    break;
                case "$defs":
                case "definitions":
                    // Compile definitions too, so invalid ones fail at load rather than at first use.
                    if (value.Kind == JsonDataKind.Object) {
                        foreach (KeyValuePair<string, JsonData> entry in value.Properties) {
                            if (entry.Value.Kind is JsonDataKind.Object or JsonDataKind.Boolean) Compile(entry.Value, resource);
                        }
                    }
                    break;
            }
        }
        if (node.Properties is { Length: > 8 }) {
            node.PropertyLookup = new Dictionary<string, JsonSchemaNode>(node.Properties.Length, StringComparer.Ordinal);
            foreach (KeyValuePair<string, JsonSchemaNode> entry in node.Properties) node.PropertyLookup[entry.Key] = entry.Value;
        }
        return node;
    }

    private void ReadDependencies(JsonSchemaNode node, JsonData value, JsonSchemaResource resource) {
        if (value.Kind != JsonDataKind.Object) throw Invalid(value, "'dependencies' must be an object");
        var required = new List<KeyValuePair<string, string[]>>();
        var schemas = new List<KeyValuePair<string, JsonSchemaNode>>();
        foreach (KeyValuePair<string, JsonData> entry in value.Properties) {
            if (entry.Value.Kind == JsonDataKind.Array) required.Add(new KeyValuePair<string, string[]>(entry.Key, ReadStrings(entry.Value, "dependencies")));
            else schemas.Add(new KeyValuePair<string, JsonSchemaNode>(entry.Key, Compile(entry.Value, resource)));
        }
        if (required.Count != 0) {
            node.DependentRequired = required.ToArray();
            node.DependentRequiredKeyword = "dependencies";
        }
        if (schemas.Count != 0) {
            node.DependentSchemas = schemas.ToArray();
            node.DependentSchemasKeyword = "dependencies";
        }
    }

    private JsonSchemaNode[] CompileArray(JsonData value, JsonSchemaResource resource, string keyword) {
        if (value.Kind != JsonDataKind.Array || value.Items.Count == 0) throw Invalid(value, "'" + keyword + "' must be a non-empty array of schemas");
        var nodes = new JsonSchemaNode[value.Items.Count];
        for (int i = 0; i < nodes.Length; i++) nodes[i] = Compile(value.Items[i], resource);
        return nodes;
    }

    private KeyValuePair<string, JsonSchemaNode>[] CompileMap(JsonData value, JsonSchemaResource resource, string keyword) {
        if (value.Kind != JsonDataKind.Object) throw Invalid(value, "'" + keyword + "' must be an object");
        var nodes = new KeyValuePair<string, JsonSchemaNode>[value.Properties.Count];
        for (int i = 0; i < nodes.Length; i++) {
            KeyValuePair<string, JsonData> entry = value.Properties[i];
            nodes[i] = new KeyValuePair<string, JsonSchemaNode>(entry.Key, Compile(entry.Value, resource));
        }
        return nodes;
    }

    private static JsonSchemaTypes ReadTypes(JsonData value) {
        if (value.Kind == JsonDataKind.String) return ReadType(value);
        if (value.Kind != JsonDataKind.Array) throw Invalid(value, "'type' must be a string or an array of strings");
        JsonSchemaTypes types = JsonSchemaTypes.None;
        foreach (JsonData item in value.Items) types |= ReadType(item);
        return types;
    }

    private static JsonSchemaTypes ReadType(JsonData value) {
        if (value.Kind != JsonDataKind.String) throw Invalid(value, "'type' must name types as strings");
        return value.GetString() switch {
            "null" => JsonSchemaTypes.Null,
            "boolean" => JsonSchemaTypes.Boolean,
            "object" => JsonSchemaTypes.Object,
            "array" => JsonSchemaTypes.Array,
            "number" => JsonSchemaTypes.Number,
            "string" => JsonSchemaTypes.String,
            "integer" => JsonSchemaTypes.Integer,
            _ => throw Invalid(value, "'" + value.GetString() + "' is not a JSON Schema type")
        };
    }

    private static JsonData ReadNumber(JsonData value, string keyword) =>
        value.Kind == JsonDataKind.Number ? value : throw Invalid(value, "'" + keyword + "' must be a number");

    private static long ReadCount(JsonData value, string keyword) {
        if (value.Kind != JsonDataKind.Number || !value.IsInteger || value.DoubleValue < 0) throw Invalid(value, "'" + keyword + "' must be a non-negative integer");
        double d = value.DoubleValue;
        return d >= long.MaxValue ? long.MaxValue : (long)d;
    }

    private static bool ReadBoolean(JsonData value, string keyword) =>
        value.Kind == JsonDataKind.Boolean ? value.GetBoolean() : throw Invalid(value, "'" + keyword + "' must be a boolean");

    private static string ReadString(JsonData value, string keyword) =>
        value.Kind == JsonDataKind.String ? value.GetString() : throw Invalid(value, "'" + keyword + "' must be a string");

    private static string[] ReadStrings(JsonData value, string keyword) {
        if (value.Kind != JsonDataKind.Array) throw Invalid(value, "'" + keyword + "' must be an array of strings");
        var names = new string[value.Items.Count];
        for (int i = 0; i < names.Length; i++) names[i] = ReadString(value.Items[i], keyword);
        return names;
    }

    private static KeyValuePair<string, string[]>[] ReadDependentRequired(JsonData value, string keyword) {
        if (value.Kind != JsonDataKind.Object) throw Invalid(value, "'" + keyword + "' must be an object");
        var entries = new KeyValuePair<string, string[]>[value.Properties.Count];
        for (int i = 0; i < entries.Length; i++) entries[i] = new KeyValuePair<string, string[]>(value.Properties[i].Key, ReadStrings(value.Properties[i].Value, keyword));
        return entries;
    }

    private Regex CreateRegex(string pattern, JsonData at) {
        try {
            return JsonSchemaRegex.Create(pattern, _options.RegexTimeout);
        } catch (ArgumentException ex) {
            throw new JsonSchemaException(Location(at) + "'" + pattern + "' is not a valid ECMA-262 regular expression: " + ex.Message, ex);
        }
    }

    private static JsonSchemaException Invalid(JsonData at, string message) => new JsonSchemaException(Location(at) + message + ".");

    private static string Location(JsonData at) =>
        at.Line > 0 ? "Schema line " + at.Line.ToString(CultureInfo.InvariantCulture) + ", column " + at.Column.ToString(CultureInfo.InvariantCulture) + ": " : string.Empty;

    #endregion

    #region References

    /// <summary>Resolves <c>$ref</c> or <c>$dynamicRef</c> against the schema's resource; returns null with a message when it cannot.</summary>
    private JsonSchemaNode? ResolveReference(JsonSchemaNode node, string reference, out string? error) {
        lock (_sync) {
            string target = JsonSchemaUri.Resolve(node.Resource.Uri, reference);
            string resourceUri = JsonSchemaUri.ResourcePart(target);
            JsonSchemaResource? resource = FindResource(resourceUri);
            if (resource is null) {
                error = _options.ReferenceResolver is null
                    ? "Cannot resolve the reference '" + reference + "': '" + resourceUri + "' is not part of this schema; set JsonSchemaOptions.ReferenceResolver to supply it."
                    : "Cannot resolve the reference '" + reference + "': the reference resolver returned no schema for '" + resourceUri + "'.";
                return null;
            }
            string fragment = JsonSchemaUri.Fragment(target);
            JsonData? data;
            if (fragment.Length == 0) {
                data = resource.Root;
            } else if (fragment[0] == '/') {
                data = ResolvePointer(resource.Root, fragment);
            } else {
                data = resource.Anchors.TryGetValue(fragment, out JsonData? anchored) ? anchored : null;
            }
            if (data is null || data.Kind is not (JsonDataKind.Object or JsonDataKind.Boolean)) {
                error = "Cannot resolve the reference '" + reference + "': '" + target + "' does not name a schema.";
                return null;
            }
            error = null;
            return Compile(data, resource);
        }
    }

    /// <summary>Finds the outermost dynamic anchor with the name in the dynamic scope, or null.</summary>
    private JsonSchemaNode? ResolveDynamicAnchor(string name, ScopeEntry? scope) {
        JsonData? found = null;
        JsonSchemaResource? foundResource = null;
        for (ScopeEntry? entry = scope; entry is not null; entry = entry.Parent) {
            if (entry.Resource.DynamicAnchors is not null && entry.Resource.DynamicAnchors.TryGetValue(name, out JsonData? data)) {
                found = data;
                foundResource = entry.Resource;
            }
        }
        if (found is null) return null;
        lock (_sync) return Compile(found, foundResource!);
    }

    private JsonSchemaResource? FindResource(string uri) {
        if (_resources.TryGetValue(uri, out JsonSchemaResource? resource)) return resource;
        Func<string, JsonData?>? resolver = _options.ReferenceResolver;
        JsonData? data = resolver?.Invoke(uri);
        if (data is null || data.Kind is not (JsonDataKind.Object or JsonDataKind.Boolean)) {
            _resources[uri] = null;
            return null;
        }
        resource = new JsonSchemaResource(uri, data);
        _resources[uri] = resource;
        Discover(data, resource, 0);
        return _resources[uri];
    }

    private static JsonData? ResolvePointer(JsonData root, string pointer) {
        if (!JsonPointerPath.TryParse(pointer, out List<string> tokens)) return null;
        JsonData current = root;
        foreach (string token in tokens) {
            if (current.Kind == JsonDataKind.Object) {
                if (!current.TryGetProperty(token, out current)) return null;
            } else if (current.Kind == JsonDataKind.Array) {
                if (!int.TryParse(token, NumberStyles.None, CultureInfo.InvariantCulture, out int index) || index >= current.Items.Count
                    || (token.Length > 1 && token[0] == '0')) return null;
                current = current.Items[index];
            } else {
                return null;
            }
        }
        return current;
    }

    #endregion
}
