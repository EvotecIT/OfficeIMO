using System.Collections.Immutable;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Text;

namespace OfficeIMO.Serialization.SourceGeneration;

/// <summary>The output for one context: its source, or the diagnostics that stopped it. Equatable for incremental caching.</summary>
internal sealed class SerializerGenerationResult : IEquatable<SerializerGenerationResult> {
    public SerializerGenerationResult(string hintName, string? source, ImmutableArray<SerializerDiagnosticInfo> diagnostics) {
        HintName = hintName;
        Source = source;
        Diagnostics = diagnostics;
    }

    public string HintName { get; }

    public string? Source { get; }

    public ImmutableArray<SerializerDiagnosticInfo> Diagnostics { get; }

    public bool Equals(SerializerGenerationResult? other) =>
        other is not null && HintName == other.HintName && Source == other.Source && Diagnostics.SequenceEqual(other.Diagnostics);

    public override bool Equals(object? obj) => Equals(obj as SerializerGenerationResult);

    public override int GetHashCode() => (HintName.GetHashCode() * 397) ^ (Source?.GetHashCode() ?? 0);
}

/// <summary>A diagnostic captured without symbols, so results stay comparable across generator runs.</summary>
internal sealed class SerializerDiagnosticInfo : IEquatable<SerializerDiagnosticInfo> {
    private readonly DiagnosticDescriptor _descriptor;
    private readonly string _path;
    private readonly TextSpan _span;
    private readonly LinePositionSpan _lines;
    private readonly string[] _arguments;

    private SerializerDiagnosticInfo(DiagnosticDescriptor descriptor, string path, TextSpan span, LinePositionSpan lines, string[] arguments) {
        _descriptor = descriptor;
        _path = path;
        _span = span;
        _lines = lines;
        _arguments = arguments;
    }

    public static SerializerDiagnosticInfo From(DiagnosticDescriptor descriptor, Location? location, params string[] arguments) {
        if (location is null || !location.IsInSource) return new SerializerDiagnosticInfo(descriptor, string.Empty, default, default, arguments);
        FileLinePositionSpan lines = location.GetLineSpan();
        return new SerializerDiagnosticInfo(descriptor, lines.Path ?? string.Empty, location.SourceSpan, lines.Span, arguments);
    }

    public string Id => _descriptor.Id;

    public Diagnostic Create() {
        Location location = _path.Length == 0 ? Location.None : Location.Create(_path, _span, _lines);
        return Diagnostic.Create(_descriptor, location, _arguments);
    }

    public bool Equals(SerializerDiagnosticInfo? other) =>
        other is not null && ReferenceEquals(_descriptor, other._descriptor) && _path == other._path && _span == other._span && _arguments.SequenceEqual(other._arguments);

    public override bool Equals(object? obj) => Equals(obj as SerializerDiagnosticInfo);

    public override int GetHashCode() => (_descriptor.Id.GetHashCode() * 397) ^ _span.GetHashCode();
}

/// <summary>Registers one format's generator pipeline: every class carrying the format's serializable attribute.</summary>
internal static class SerializerGeneratorPipeline {
    public static void Initialize(IncrementalGeneratorInitializationContext context, SerializerFormat format) {
        IncrementalValuesProvider<SerializerGenerationResult> results = context.SyntaxProvider.ForAttributeWithMetadataName(
            format.SerializableAttribute,
            static (node, _) => node is Microsoft.CodeAnalysis.CSharp.Syntax.ClassDeclarationSyntax,
            (attributeContext, cancellationToken) => SerializerContextParser.Parse(format, attributeContext, cancellationToken));

        context.RegisterSourceOutput(results, static (production, result) => {
            foreach (SerializerDiagnosticInfo diagnostic in result.Diagnostics) production.ReportDiagnostic(diagnostic.Create());
            if (result.Source is not null) production.AddSource(result.HintName, SourceText.From(result.Source, System.Text.Encoding.UTF8));
        });
    }
}
