using Microsoft.CodeAnalysis;

namespace OfficeIMO.Serialization.SourceGeneration;

/// <summary>
/// What differs between the OfficeIMO serializers' source generators. The shape rules, the shared attributes in
/// <c>OfficeIMO</c> and the emitted metadata calls are the same; each format names its own context, metadata types and
/// diagnostics, which each generator declares so analyzer release tracking sees them. The metadata API of a format is
/// <c>{Namespace}.{Prefix}MetadataServices</c>, <c>{Prefix}TypeInfo&lt;T&gt;</c>, <c>{Prefix}ObjectInfoValues&lt;T&gt;</c>
/// and so on.
/// </summary>
internal sealed class SerializerFormat {
    public SerializerFormat(string prefix, string displayName, string? inlineTableAttribute, IEnumerable<string> extraBuiltInScalars,
        DiagnosticDescriptor contextNotPartial, DiagnosticDescriptor unsupportedType, DiagnosticDescriptor noConstructor,
        DiagnosticDescriptor duplicateName, DiagnosticDescriptor invalidDerivedType) {
        Prefix = prefix;
        DisplayName = displayName;
        Namespace = "OfficeIMO." + prefix;
        ContextTypeName = Namespace + "." + prefix + "SerializerContext";
        SerializableAttribute = Namespace + "." + prefix + "SerializableAttribute";
        InlineTableAttribute = inlineTableAttribute;
        ExtraBuiltInScalars = new HashSet<string>(extraBuiltInScalars, StringComparer.Ordinal);
        ContextNotPartial = contextNotPartial;
        UnsupportedType = unsupportedType;
        NoConstructor = noConstructor;
        DuplicateName = duplicateName;
        InvalidDerivedType = invalidDerivedType;
    }

    /// <summary>The type-name prefix, <c>Yaml</c> or <c>Toml</c>.</summary>
    public string Prefix { get; }

    /// <summary>The format's name in messages, <c>YAML</c> or <c>TOML</c>.</summary>
    public string DisplayName { get; }

    public string Namespace { get; }

    public string ContextTypeName { get; }

    public string SerializableAttribute { get; }

    /// <summary>The full name of an attribute that writes a type or property as an inline table, or null.</summary>
    public string? InlineTableAttribute { get; }

    /// <summary>Format types the runtime handles as built-in scalars, beyond the shared set.</summary>
    public HashSet<string> ExtraBuiltInScalars { get; }

    public DiagnosticDescriptor ContextNotPartial { get; }

    public DiagnosticDescriptor UnsupportedType { get; }

    public DiagnosticDescriptor NoConstructor { get; }

    public DiagnosticDescriptor DuplicateName { get; }

    public DiagnosticDescriptor InvalidDerivedType { get; }

    /// <summary>A type of the format's namespace, written fully qualified for generated code.</summary>
    public string Global(string name) => "global::" + Namespace + "." + Prefix + name;
}
