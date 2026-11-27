using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Text;

namespace OfficeIMO.Fluent.Generators {
    /// <summary>The build-time diagnostics. Every one is an error: a skipped method would silently change the fluent surface.</summary>
    internal static class FluentDiagnostics {
        private const string Category = "OfficeIMO.Fluent";

        internal static readonly DiagnosticDescriptor Collision = new DiagnosticDescriptor("OIFLUENT001",
            "A generated fluent method would collide with an existing method",
            "The fluent method '{0}' is not generated: it would collide with {1}", Category, DiagnosticSeverity.Error, true);

        internal static readonly DiagnosticDescriptor ValueTypeSetter = new DiagnosticDescriptor("OIFLUENT002",
            "A value-type property setter cannot be generated as a fluent extension",
            "Property '{0}' on value type '{1}' would be set on a copy; mark the handle's write path with [FluentMutator] or the property with [NoFluent]",
            Category, DiagnosticSeverity.Error, true);

        internal static readonly DiagnosticDescriptor InvalidMutator = new DiagnosticDescriptor("OIFLUENT003",
            "A fluent mutator must be an accessible instance method",
            "Method '{0}' cannot be a fluent mutator: it must be a public or internal non-generic instance method without ref or out parameters",
            Category, DiagnosticSeverity.Error, true);

        internal static readonly DiagnosticDescriptor UnsupportedType = new DiagnosticDescriptor("OIFLUENT004",
            "Fluent generation requires a top-level non-generic class or struct",
            "Type '{0}' cannot be marked [Fluent]: use a top-level, non-generic, non-static class or struct",
            Category, DiagnosticSeverity.Error, true);

        internal static readonly DiagnosticDescriptor Duplicate = new DiagnosticDescriptor("OIFLUENT005",
            "Two members map to the same fluent method",
            "More than one member maps to the fluent method '{0}'", Category, DiagnosticSeverity.Error, true);
    }

    /// <summary>A diagnostic kept as plain values, so the incremental pipeline can compare models.</summary>
    internal sealed class DiagnosticInfo : IEquatable<DiagnosticInfo> {
        private DiagnosticInfo(DiagnosticDescriptor descriptor, string? path, TextSpan span, LinePositionSpan lines, string[] arguments) {
            Descriptor = descriptor;
            Path = path;
            Span = span;
            Lines = lines;
            Arguments = arguments;
        }

        internal DiagnosticDescriptor Descriptor { get; }
        internal string? Path { get; }
        internal TextSpan Span { get; }
        internal LinePositionSpan Lines { get; }
        internal string[] Arguments { get; }

        internal static DiagnosticInfo From(DiagnosticDescriptor descriptor, ISymbol symbol, params string[] arguments) {
            var location = symbol.Locations.FirstOrDefault(l => l.IsInSource);
            return location == null
                ? new DiagnosticInfo(descriptor, null, default, default, arguments)
                : new DiagnosticInfo(descriptor, location.SourceTree?.FilePath, location.SourceSpan, location.GetLineSpan().Span, arguments);
        }

        internal Diagnostic Create() => Diagnostic.Create(Descriptor,
            Path == null ? Location.None : Location.Create(Path, Span, Lines), Arguments.Cast<object>().ToArray());

        public bool Equals(DiagnosticInfo? other) => other != null && Descriptor.Id == other.Descriptor.Id && Path == other.Path
            && Span.Equals(other.Span) && Arguments.SequenceEqual(other.Arguments);

        public override bool Equals(object? obj) => Equals(obj as DiagnosticInfo);

        public override int GetHashCode() => (Descriptor.Id, Path, Span).GetHashCode();
    }

    /// <summary>The generated source and diagnostics of one marked type.</summary>
    internal sealed class TypeModel : IEquatable<TypeModel> {
        internal TypeModel(string hintName, string? source, IReadOnlyList<DiagnosticInfo> diagnostics) {
            HintName = hintName;
            Source = source;
            Diagnostics = diagnostics;
        }

        internal string HintName { get; }
        internal string? Source { get; }
        internal IReadOnlyList<DiagnosticInfo> Diagnostics { get; }

        public bool Equals(TypeModel? other) => other != null && HintName == other.HintName && Source == other.Source && Diagnostics.SequenceEqual(other.Diagnostics);

        public override bool Equals(object? obj) => Equals(obj as TypeModel);

        public override int GetHashCode() => (HintName, Source).GetHashCode();
    }
}
