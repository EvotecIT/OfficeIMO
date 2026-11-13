using Microsoft.CodeAnalysis;
using OfficeIMO.Serialization.SourceGeneration;

namespace OfficeIMO.Yaml.SourceGeneration;

/// <summary>
/// Generates serialization metadata for each partial class that derives from <c>OfficeIMO.Yaml.YamlSerializerContext</c>
/// and lists root types with <c>[YamlSerializable]</c>. Every type the roots reach is classified with the same shape
/// rules as the runtime's reflection metadata; a shape the serializer cannot handle is a build error, never a runtime
/// fallback. The shape rules and the emitter are shared with the TOML generator.
/// </summary>
[Generator(LanguageNames.CSharp)]
public sealed class YamlSourceGenerator : IIncrementalGenerator {
    private static readonly SerializerFormat Format = new SerializerFormat("Yaml", "YAML", inlineTableAttribute: null, Array.Empty<string>(),
        YamlDiagnostics.ContextNotPartial, YamlDiagnostics.UnsupportedType, YamlDiagnostics.NoConstructor, YamlDiagnostics.DuplicateName, YamlDiagnostics.InvalidDerivedType);

    /// <inheritdoc />
    public void Initialize(IncrementalGeneratorInitializationContext context) => SerializerGeneratorPipeline.Initialize(context, Format);
}

internal static class YamlDiagnostics {
    private const string Category = "OfficeIMO.Yaml";

    public static readonly DiagnosticDescriptor ContextNotPartial = new(
        "YAML001",
        "A YAML serializer context must be a partial, non-generic class",
        "The YAML serializer context '{0}' must be a partial, non-generic class, and every type that contains it must be partial",
        Category, DiagnosticSeverity.Error, isEnabledByDefault: true);

    public static readonly DiagnosticDescriptor UnsupportedType = new(
        "YAML002",
        "The YAML serializer does not support this type",
        "The type '{0}' cannot be serialized: {1}",
        Category, DiagnosticSeverity.Error, isEnabledByDefault: true);

    public static readonly DiagnosticDescriptor NoConstructor = new(
        "YAML003",
        "The type has no constructor reading can use",
        "The type '{0}' cannot be read: {1}",
        Category, DiagnosticSeverity.Error, isEnabledByDefault: true);

    public static readonly DiagnosticDescriptor DuplicateName = new(
        "YAML004",
        "Two properties have the same YAML name",
        "The type '{0}' has more than one property named '{1}' in YAML",
        Category, DiagnosticSeverity.Error, isEnabledByDefault: true);

    public static readonly DiagnosticDescriptor InvalidDerivedType = new(
        "YAML005",
        "The derived-type registration is invalid",
        "The derived-type registration on '{0}' is invalid: {1}",
        Category, DiagnosticSeverity.Error, isEnabledByDefault: true);
}
