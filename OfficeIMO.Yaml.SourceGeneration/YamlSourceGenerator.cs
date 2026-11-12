using System.Collections.Immutable;
using System.Text;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.Text;

namespace OfficeIMO.Yaml.SourceGeneration;

/// <summary>
/// Generates serialization metadata for each partial class that derives from <c>OfficeIMO.Yaml.YamlSerializerContext</c>
/// and lists root types with <c>[YamlSerializable]</c>. Every type the roots reach is classified with the same shape
/// rules as the runtime's reflection metadata; a shape the serializer cannot handle is a build error, never a runtime
/// fallback.
/// </summary>
[Generator(LanguageNames.CSharp)]
public sealed class YamlSourceGenerator : IIncrementalGenerator {
    internal const string SerializableAttributeName = "OfficeIMO.Yaml.YamlSerializableAttribute";

    /// <inheritdoc />
    public void Initialize(IncrementalGeneratorInitializationContext context) {
        IncrementalValuesProvider<YamlGenerationResult> results = context.SyntaxProvider.ForAttributeWithMetadataName(
            SerializableAttributeName,
            static (node, _) => node is ClassDeclarationSyntax,
            static (attributeContext, cancellationToken) => YamlContextParser.Parse(attributeContext, cancellationToken));

        context.RegisterSourceOutput(results, static (production, result) => {
            foreach (YamlDiagnosticInfo diagnostic in result.Diagnostics) production.ReportDiagnostic(diagnostic.Create());
            if (result.Source is not null) production.AddSource(result.HintName, SourceText.From(result.Source, Encoding.UTF8));
        });
    }
}

/// <summary>The output for one context: its source, or the diagnostics that stopped it. Equatable for incremental caching.</summary>
internal sealed class YamlGenerationResult : IEquatable<YamlGenerationResult> {
    public YamlGenerationResult(string hintName, string? source, ImmutableArray<YamlDiagnosticInfo> diagnostics) {
        HintName = hintName;
        Source = source;
        Diagnostics = diagnostics;
    }

    public string HintName { get; }

    public string? Source { get; }

    public ImmutableArray<YamlDiagnosticInfo> Diagnostics { get; }

    public bool Equals(YamlGenerationResult? other) =>
        other is not null && HintName == other.HintName && Source == other.Source && Diagnostics.SequenceEqual(other.Diagnostics);

    public override bool Equals(object? obj) => Equals(obj as YamlGenerationResult);

    public override int GetHashCode() => (HintName.GetHashCode() * 397) ^ (Source?.GetHashCode() ?? 0);
}

/// <summary>A diagnostic captured without symbols, so results stay comparable across generator runs.</summary>
internal sealed class YamlDiagnosticInfo : IEquatable<YamlDiagnosticInfo> {
    private readonly DiagnosticDescriptor _descriptor;
    private readonly string _path;
    private readonly TextSpan _span;
    private readonly LinePositionSpan _lines;
    private readonly string[] _arguments;

    private YamlDiagnosticInfo(DiagnosticDescriptor descriptor, string path, TextSpan span, LinePositionSpan lines, string[] arguments) {
        _descriptor = descriptor;
        _path = path;
        _span = span;
        _lines = lines;
        _arguments = arguments;
    }

    public static YamlDiagnosticInfo From(DiagnosticDescriptor descriptor, Location? location, params string[] arguments) {
        if (location is null || !location.IsInSource) return new YamlDiagnosticInfo(descriptor, string.Empty, default, default, arguments);
        FileLinePositionSpan lines = location.GetLineSpan();
        return new YamlDiagnosticInfo(descriptor, lines.Path ?? string.Empty, location.SourceSpan, lines.Span, arguments);
    }

    public string Id => _descriptor.Id;

    public Diagnostic Create() {
        Location location = _path.Length == 0 ? Location.None : Location.Create(_path, _span, _lines);
        return Diagnostic.Create(_descriptor, location, _arguments);
    }

    public bool Equals(YamlDiagnosticInfo? other) =>
        other is not null && ReferenceEquals(_descriptor, other._descriptor) && _path == other._path && _span == other._span && _arguments.SequenceEqual(other._arguments);

    public override bool Equals(object? obj) => Equals(obj as YamlDiagnosticInfo);

    public override int GetHashCode() => (_descriptor.Id.GetHashCode() * 397) ^ _span.GetHashCode();
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
