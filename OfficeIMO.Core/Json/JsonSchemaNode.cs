using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeIMO.Json;

/// <summary>A schema resource: a document or an embedded schema with its own <c>$id</c>, and the anchors inside it.</summary>
internal sealed class JsonSchemaResource {
    public JsonSchemaResource(string uri, JsonData root) {
        Uri = uri;
        Root = root;
    }

    /// <summary>The absolute URI without fragment, or a relative or empty one when no base is known.</summary>
    public string Uri { get; }

    public JsonData Root { get; }

    public Dictionary<string, JsonData> Anchors { get; } = new Dictionary<string, JsonData>(StringComparer.Ordinal);

    public Dictionary<string, JsonData>? DynamicAnchors { get; set; }
}

[Flags]
internal enum JsonSchemaTypes : byte {
    None = 0,
    Null = 1,
    Boolean = 2,
    Object = 4,
    Array = 8,
    Number = 16,
    String = 32,
    Integer = 64
}

/// <summary>A compiled schema object or boolean schema. Subschemas are compiled eagerly; references resolve on first use.</summary>
internal sealed class JsonSchemaNode {
    public JsonSchemaNode(JsonData data, JsonSchemaResource resource) {
        Data = data;
        Resource = resource;
    }

    public JsonData Data { get; }
    public JsonSchemaResource Resource { get; }

    /// <summary>True or false for a boolean schema; null for a schema object.</summary>
    public bool? Boolean;

    public JsonSchemaTypes Types;
    public JsonData[]? Enum;
    public JsonData? Const;
    public JsonData? MultipleOf, Maximum, ExclusiveMaximum, Minimum, ExclusiveMinimum;
    public long MaxLength = -1, MinLength = -1, MaxItems = -1, MinItems = -1, MaxProperties = -1, MinProperties = -1, MaxContains = -1, MinContains = -1;
    public Regex? Pattern;
    public string? PatternText;
    public string? Format;
    public bool UniqueItems;

    public JsonSchemaNode? Items, Contains, UnevaluatedItems;
    public JsonSchemaNode[]? PrefixItems;
    public KeyValuePair<string, JsonSchemaNode>[]? Properties;
    public Dictionary<string, JsonSchemaNode>? PropertyLookup;
    public PatternSchema[]? PatternProperties;
    public JsonSchemaNode? AdditionalProperties, PropertyNames, UnevaluatedProperties;
    public string[]? Required;
    public KeyValuePair<string, string[]>[]? DependentRequired;
    public KeyValuePair<string, JsonSchemaNode>[]? DependentSchemas;
    public JsonSchemaNode[]? AllOf, AnyOf, OneOf;
    public JsonSchemaNode? Not, If, Then, Else;

    public string? Ref, DynamicRef;
    /// <summary>Draft 7: <c>$ref</c> makes the schema ignore its other keywords.</summary>
    public bool RefOnly;
    /// <summary>
    /// The keywords the item and dependency schemas came from, for schema paths. Draft 7's array <c>items</c> and
    /// <c>additionalItems</c> map to <see cref="PrefixItems"/> and <see cref="Items"/>, and its <c>dependencies</c> to
    /// <see cref="DependentRequired"/> and <see cref="DependentSchemas"/>.
    /// </summary>
    public string PrefixItemsKeyword = "prefixItems", ItemsKeyword = "items", DependentRequiredKeyword = "dependentRequired", DependentSchemasKeyword = "dependentSchemas";

    // Reference targets resolved on first use; a failed resolution keeps its message.
    public JsonSchemaNode? RefTarget, DynamicRefTarget;
    public string? RefError, DynamicRefError;
    public string? DynamicAnchorName;
    public bool RefResolved, DynamicRefResolved;

    internal readonly struct PatternSchema {
        public PatternSchema(string pattern, Regex regex, JsonSchemaNode schema) {
            Pattern = pattern;
            Regex = regex;
            Schema = schema;
        }

        public string Pattern { get; }
        public Regex Regex { get; }
        public JsonSchemaNode Schema { get; }
    }
}
