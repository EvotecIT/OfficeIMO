using Microsoft.CodeAnalysis;
using OfficeIMO.Serialization.SourceGeneration;

namespace OfficeIMO.Toml.SourceGeneration;

/// <summary>
/// Generates serialization metadata for each partial class that derives from <c>OfficeIMO.Toml.TomlSerializerContext</c>
/// and lists root types with <c>[TomlSerializable]</c>. Every type the roots reach is classified with the same shape
/// rules as the runtime's reflection metadata; a shape the serializer cannot handle is a build error, never a runtime
/// fallback. The shape rules and the emitter are shared with the YAML generator; <c>[TomlInlineTable]</c> is the one
/// TOML-only attribute it reads.
/// </summary>
[Generator(LanguageNames.CSharp)]
public sealed class TomlSourceGenerator : IIncrementalGenerator {
    private static readonly SerializerFormat Format = new SerializerFormat("Toml", "TOML", "OfficeIMO.Toml.TomlInlineTableAttribute", new[] { "OfficeIMO.Toml.TomlDateTime" },
        TomlDiagnostics.ContextNotPartial, TomlDiagnostics.UnsupportedType, TomlDiagnostics.NoConstructor, TomlDiagnostics.DuplicateName, TomlDiagnostics.InvalidDerivedType);

    /// <inheritdoc />
    public void Initialize(IncrementalGeneratorInitializationContext context) => SerializerGeneratorPipeline.Initialize(context, Format);
}

internal static class TomlDiagnostics {
    private const string Category = "OfficeIMO.Toml";

    public static readonly DiagnosticDescriptor ContextNotPartial = new(
        "TOML001",
        "A TOML serializer context must be a partial, non-generic class",
        "The TOML serializer context '{0}' must be a partial, non-generic class, and every type that contains it must be partial",
        Category, DiagnosticSeverity.Error, isEnabledByDefault: true);

    public static readonly DiagnosticDescriptor UnsupportedType = new(
        "TOML002",
        "The TOML serializer does not support this type",
        "The type '{0}' cannot be serialized: {1}",
        Category, DiagnosticSeverity.Error, isEnabledByDefault: true);

    public static readonly DiagnosticDescriptor NoConstructor = new(
        "TOML003",
        "The type has no constructor reading can use",
        "The type '{0}' cannot be read: {1}",
        Category, DiagnosticSeverity.Error, isEnabledByDefault: true);

    public static readonly DiagnosticDescriptor DuplicateName = new(
        "TOML004",
        "Two properties have the same TOML name",
        "The type '{0}' has more than one property named '{1}' in TOML",
        Category, DiagnosticSeverity.Error, isEnabledByDefault: true);

    public static readonly DiagnosticDescriptor InvalidDerivedType = new(
        "TOML005",
        "The derived-type registration is invalid",
        "The derived-type registration on '{0}' is invalid: {1}",
        Category, DiagnosticSeverity.Error, isEnabledByDefault: true);
}
