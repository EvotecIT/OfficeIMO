using System.Collections.Immutable;
using System.Globalization;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;

namespace OfficeIMO.Yaml.SourceGeneration;

internal enum YamlShape {
    BuiltIn,
    Nullable,
    Enum,
    Array,
    List,
    Set,
    CustomCollection,
    Dictionary,
    CustomDictionary,
    Object
}

internal sealed class YamlTypeEntry {
    public YamlTypeEntry(ITypeSymbol type, YamlShape shape) {
        Type = type;
        Shape = shape;
        FullName = type.ToDisplayString(SymbolDisplayFormat.FullyQualifiedFormat);
    }

    public ITypeSymbol Type { get; }
    public YamlShape Shape { get; }
    public string FullName { get; }
    public string PropertyName { get; set; } = string.Empty;
    public ITypeSymbol? Element { get; set; }
    public ITypeSymbol? Key { get; set; }
    public ITypeSymbol? Value { get; set; }
    public YamlObjectShape? Object { get; set; }
}

internal sealed class YamlObjectShape {
    public List<YamlPropertyShape> Properties { get; } = new List<YamlPropertyShape>();
    public IMethodSymbol? Constructor { get; set; }
    public bool UseDefault { get; set; }
    public bool HasRequiredMembers { get; set; }
    public List<(ITypeSymbol Type, string Discriminator)> Derived { get; } = new List<(ITypeSymbol, string)>();
    public string? DiscriminatorPropertyName { get; set; }
}

internal enum YamlSetter {
    None,
    Direct,
    Accessor,
    Reflection
}

internal sealed class YamlPropertyShape {
    public YamlPropertyShape(IPropertySymbol symbol) {
        Symbol = symbol;
    }

    public IPropertySymbol Symbol { get; }
    public string? YamlName { get; set; }
    public bool IsRequired { get; set; }
    public string? IgnoreCondition { get; set; }
    public YamlSetter Setter { get; set; }
}

/// <summary>
/// Walks the type graph of one context and classifies every type with the runtime's shape rules
/// (<c>YamlTypeShapes</c> in OfficeIMO.Yaml): built-in scalars, nullable values, enums, arrays, lists and sets,
/// dictionaries, and objects with their properties, constructor and derived types.
/// </summary>
internal sealed class YamlContextParser {
    private readonly INamedTypeSymbol _context;
    private readonly Location? _location;
    private readonly List<YamlDiagnosticInfo> _diagnostics = new List<YamlDiagnosticInfo>();
    private readonly Dictionary<ITypeSymbol, YamlTypeEntry> _entries = new Dictionary<ITypeSymbol, YamlTypeEntry>(SymbolEqualityComparer.Default);
    private readonly List<YamlTypeEntry> _order = new List<YamlTypeEntry>();
    private readonly bool _unsafeAccessor;
    private readonly bool _genericUnsafeAccessor;

    private YamlContextParser(Compilation compilation, INamedTypeSymbol context, Location? location) {
        _context = context;
        _location = location;
        _unsafeAccessor = compilation.GetTypeByMetadataName("System.Runtime.CompilerServices.UnsafeAccessorAttribute") is not null;
        // Generic declaring types work with UnsafeAccessor from .NET 9.
        _genericUnsafeAccessor = _unsafeAccessor && compilation.GetSpecialType(SpecialType.System_Object).ContainingAssembly.Identity.Version.Major >= 9;
    }

    public static YamlGenerationResult Parse(GeneratorAttributeSyntaxContext attributeContext, CancellationToken cancellationToken) {
        var context = (INamedTypeSymbol)attributeContext.TargetSymbol;
        Location? location = attributeContext.TargetNode.GetLocation();
        string hintName = HintName(context);
        var parser = new YamlContextParser(attributeContext.SemanticModel.Compilation, context, location);

        if (!parser.IsValidContext(context)) {
            return new YamlGenerationResult(hintName, null, parser._diagnostics.ToImmutableArray());
        }

        foreach (AttributeData attribute in attributeContext.Attributes) {
            cancellationToken.ThrowIfCancellationRequested();
            if (attribute.ConstructorArguments.Length == 1 && attribute.ConstructorArguments[0].Value is ITypeSymbol root) {
                Location? rootLocation = attribute.ApplicationSyntaxReference?.GetSyntax(cancellationToken).GetLocation() ?? location;
                parser.Add(root, rootLocation, isRoot: true);
            }
        }
        // Walk breadth-first; entries added while walking are walked too.
        for (int i = 0; i < parser._order.Count; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            parser.Expand(parser._order[i]);
        }

        if (parser._diagnostics.Count > 0) return new YamlGenerationResult(hintName, null, parser._diagnostics.ToImmutableArray());
        string source = YamlContextEmitter.Emit(context, parser._order, parser._roots, parser._unsafeAccessor);
        return new YamlGenerationResult(hintName, source, ImmutableArray<YamlDiagnosticInfo>.Empty);
    }

    private readonly List<YamlTypeEntry> _roots = new List<YamlTypeEntry>();

    private static string HintName(INamedTypeSymbol context) {
        string name = context.ToDisplayString(SymbolDisplayFormat.FullyQualifiedFormat).Replace("global::", string.Empty);
        var safe = new System.Text.StringBuilder(name.Length);
        foreach (char c in name) safe.Append(char.IsLetterOrDigit(c) || c == '.' || c == '_' ? c : '_');
        return safe + ".YamlContext.g.cs";
    }

    private bool IsValidContext(INamedTypeSymbol context) {
        bool valid = context.TypeParameters.Length == 0 && IsPartial(context) && DerivesFromContext(context);
        for (INamedTypeSymbol? outer = context.ContainingType; outer is not null && valid; outer = outer.ContainingType) {
            valid = IsPartial(outer) && outer.TypeParameters.Length == 0;
        }
        if (!valid) _diagnostics.Add(YamlDiagnosticInfo.From(YamlDiagnostics.ContextNotPartial, _location, context.Name));
        return valid;
    }

    private static bool IsPartial(INamedTypeSymbol type) {
        foreach (SyntaxReference reference in type.DeclaringSyntaxReferences) {
            if (reference.GetSyntax() is TypeDeclarationSyntax declaration && declaration.Modifiers.Any(SyntaxKind.PartialKeyword)) return true;
        }
        return false;
    }

    private static bool DerivesFromContext(INamedTypeSymbol type) {
        for (INamedTypeSymbol? current = type.BaseType; current is not null; current = current.BaseType) {
            if (current.ToDisplayString() == "OfficeIMO.Yaml.YamlSerializerContext") return true;
        }
        return false;
    }

    // ------------------------------------------------------------------ classification

    private YamlTypeEntry? Add(ITypeSymbol type, Location? location, bool isRoot = false) {
        type = Normalize(type);
        if (_entries.TryGetValue(type, out YamlTypeEntry? existing)) {
            if (isRoot && !_roots.Contains(existing)) _roots.Add(existing);
            return existing;
        }
        YamlTypeEntry? entry = Classify(type, location);
        if (entry is null) return null;
        _entries.Add(type, entry);
        _order.Add(entry);
        if (isRoot) _roots.Add(entry);
        return entry;
    }

    private static ITypeSymbol Normalize(ITypeSymbol type) =>
        type.IsReferenceType ? type.WithNullableAnnotation(NullableAnnotation.NotAnnotated) : type;

    private void Report(DiagnosticDescriptor descriptor, Location? location, params string[] arguments) =>
        _diagnostics.Add(YamlDiagnosticInfo.From(descriptor, location ?? _location, arguments));

    private static string Display(ITypeSymbol type) => type.ToDisplayString(SymbolDisplayFormat.CSharpErrorMessageFormat);

    private YamlTypeEntry? Classify(ITypeSymbol type, Location? location) {
        string? reason = UnsupportedReason(type);
        if (reason is not null) {
            Report(YamlDiagnostics.UnsupportedType, location, Display(type), reason);
            return null;
        }
        if (IsBuiltIn(type)) return new YamlTypeEntry(type, YamlShape.BuiltIn);

        if (type is INamedTypeSymbol { OriginalDefinition.SpecialType: SpecialType.System_Nullable_T } nullable) {
            return new YamlTypeEntry(type, YamlShape.Nullable) { Element = nullable.TypeArguments[0] };
        }
        if (type.TypeKind == TypeKind.Enum) return new YamlTypeEntry(type, YamlShape.Enum);
        if (type is IArrayTypeSymbol array) {
            if (!array.IsSZArray) {
                Report(YamlDiagnostics.UnsupportedType, location, Display(type), "only one-dimensional, zero-based arrays are supported");
                return null;
            }
            return new YamlTypeEntry(type, YamlShape.Array) { Element = array.ElementType };
        }
        var named = (INamedTypeSymbol)type;

        if (TryGetDictionary(named, out ITypeSymbol? key, out ITypeSymbol? value, out YamlShape? dictionaryShape)) {
            if (dictionaryShape is null) {
                Report(YamlDiagnostics.UnsupportedType, location, Display(type), "it implements a dictionary interface but has no public parameterless constructor");
                return null;
            }
            if (!IsKeyType(key!)) {
                Report(YamlDiagnostics.UnsupportedType, location, Display(type), "dictionary keys must be strings, numbers, enums or other scalar types");
                return null;
            }
            return new YamlTypeEntry(type, dictionaryShape.Value) { Key = key, Value = value };
        }

        if (TryGetCollection(named, out ITypeSymbol? element, out YamlShape? collectionShape)) {
            if (collectionShape is null) {
                Report(YamlDiagnostics.UnsupportedType, location, Display(type), "it is a collection without a public parameterless constructor, or implements IEnumerable<T> more than once");
                return null;
            }
            return new YamlTypeEntry(type, collectionShape.Value) { Element = element };
        }
        if (Implements(named, "System.Collections.IEnumerable")) {
            Report(YamlDiagnostics.UnsupportedType, location, Display(type), "non-generic collections are not supported");
            return null;
        }

        bool polymorphic = GetAttributes(named, "OfficeIMO.Yaml.YamlDerivedTypeAttribute").Any();
        if ((named.IsAbstract || named.TypeKind == TypeKind.Interface) && !polymorphic) {
            Report(YamlDiagnostics.UnsupportedType, location, Display(type), "it is abstract and registers no derived types with [YamlDerivedType]");
            return null;
        }
        return new YamlTypeEntry(type, YamlShape.Object);
    }

    private static string? UnsupportedReason(ITypeSymbol type) {
        switch (type.TypeKind) {
            case TypeKind.Pointer:
            case TypeKind.FunctionPointer:
                return "pointers, by-reference and ref struct types cannot be serialized";
            case TypeKind.TypeParameter:
            case TypeKind.Error:
                return "open generic types cannot be serialized";
            case TypeKind.Delegate:
                return "delegates cannot be serialized";
        }
        if (type.IsRefLikeType) return "pointers, by-reference and ref struct types cannot be serialized";
        if (type is INamedTypeSymbol { IsUnboundGenericType: true }) return "open generic types cannot be serialized";
        if (type.SpecialType is SpecialType.System_IntPtr or SpecialType.System_UIntPtr) return "the type has no YAML representation";
        for (ITypeSymbol? current = type; current is not null; current = current.BaseType) {
            if (current.ToDisplayString() == "System.Reflection.MemberInfo") return "the type has no YAML representation";
        }
        return null;
    }

    private static readonly HashSet<string> BuiltInNames = new HashSet<string>(StringComparer.Ordinal) {
        "System.Numerics.BigInteger", "System.DateTimeOffset", "System.TimeSpan", "System.Guid", "System.Uri", "System.DateOnly", "System.TimeOnly"
    };

    private static bool IsBuiltInScalar(ITypeSymbol type) {
        switch (type.SpecialType) {
            case SpecialType.System_Object:
            case SpecialType.System_Boolean:
            case SpecialType.System_Char:
            case SpecialType.System_SByte:
            case SpecialType.System_Byte:
            case SpecialType.System_Int16:
            case SpecialType.System_UInt16:
            case SpecialType.System_Int32:
            case SpecialType.System_UInt32:
            case SpecialType.System_Int64:
            case SpecialType.System_UInt64:
            case SpecialType.System_Decimal:
            case SpecialType.System_Single:
            case SpecialType.System_Double:
            case SpecialType.System_String:
            case SpecialType.System_DateTime:
                return true;
        }
        return type.TypeKind == TypeKind.Dynamic || BuiltInNames.Contains(type.ToDisplayString());
    }

    /// <summary>The types the runtime options handle without a resolver, as <c>YamlMetadataServices.TryCreateBuiltIn</c> does.</summary>
    private static bool IsBuiltIn(ITypeSymbol type) {
        if (IsBuiltInScalar(type)) return true;
        if (type is IArrayTypeSymbol { IsSZArray: true } array) {
            return array.ElementType.SpecialType is SpecialType.System_Byte or SpecialType.System_Object;
        }
        if (type is INamedTypeSymbol { IsGenericType: true } named) {
            if (named.OriginalDefinition.SpecialType == SpecialType.System_Nullable_T) return IsBuiltInScalar(named.TypeArguments[0]);
            string definition = named.OriginalDefinition.ToDisplayString();
            if (definition == "System.Collections.Generic.List<T>") return named.TypeArguments[0].SpecialType == SpecialType.System_Object;
            if (definition == "System.Collections.Generic.Dictionary<TKey, TValue>") {
                return named.TypeArguments[0].SpecialType == SpecialType.System_String && named.TypeArguments[1].SpecialType == SpecialType.System_Object;
            }
        }
        return false;
    }

    private static bool IsKeyType(ITypeSymbol type) =>
        type.TypeKind == TypeKind.Enum
        || (IsBuiltInScalar(type) && type.SpecialType != SpecialType.System_Object && type.TypeKind != TypeKind.Dynamic);

    private static string Definition(INamedTypeSymbol type) => type.OriginalDefinition.ToDisplayString();

    private static INamedTypeSymbol? FindInterface(INamedTypeSymbol type, string definition) {
        if (type.IsGenericType && Definition(type) == definition) return type;
        foreach (INamedTypeSymbol candidate in type.AllInterfaces) {
            if (candidate.IsGenericType && Definition(candidate) == definition) return candidate;
        }
        return null;
    }

    private static bool Implements(INamedTypeSymbol type, string name) {
        if (type.ToDisplayString() == name) return true;
        foreach (INamedTypeSymbol candidate in type.AllInterfaces) {
            if (candidate.ToDisplayString() == name) return true;
        }
        return false;
    }

    private static bool HasPublicParameterlessConstructor(INamedTypeSymbol type) =>
        type.IsValueType || type.InstanceConstructors.Any(static c => c.Parameters.Length == 0 && c.DeclaredAccessibility == Accessibility.Public);

    private static bool TryGetDictionary(INamedTypeSymbol type, out ITypeSymbol? key, out ITypeSymbol? value, out YamlShape? shape) {
        key = value = null;
        shape = null;
        INamedTypeSymbol? found = null;
        if (type.IsGenericType) {
            string definition = Definition(type);
            if (definition is "System.Collections.Generic.Dictionary<TKey, TValue>" or "System.Collections.Generic.IDictionary<TKey, TValue>"
                or "System.Collections.Generic.IReadOnlyDictionary<TKey, TValue>") {
                found = type;
                shape = YamlShape.Dictionary;
            }
        }
        if (found is null) {
            found = FindInterface(type, "System.Collections.Generic.IDictionary<TKey, TValue>") ?? FindInterface(type, "System.Collections.Generic.IReadOnlyDictionary<TKey, TValue>");
            if (found is null) return false;
            if (!type.IsAbstract && type.TypeKind != TypeKind.Interface && FindInterface(type, "System.Collections.Generic.IDictionary<TKey, TValue>") is not null
                && HasPublicParameterlessConstructor(type)) {
                shape = YamlShape.CustomDictionary;
            }
        }
        key = found.TypeArguments[0];
        value = found.TypeArguments[1];
        return true;
    }

    private static bool TryGetCollection(INamedTypeSymbol type, out ITypeSymbol? element, out YamlShape? shape) {
        element = null;
        shape = null;
        if (type.SpecialType == SpecialType.System_String) return false;
        if (type.IsGenericType) {
            string definition = Definition(type);
            ITypeSymbol argument = type.TypeArguments[0];
            switch (definition) {
                case "System.Collections.Generic.List<T>":
                case "System.Collections.Generic.IEnumerable<T>":
                case "System.Collections.Generic.ICollection<T>":
                case "System.Collections.Generic.IList<T>":
                case "System.Collections.Generic.IReadOnlyCollection<T>":
                case "System.Collections.Generic.IReadOnlyList<T>":
                    element = argument;
                    shape = YamlShape.List;
                    return true;
                case "System.Collections.Generic.HashSet<T>":
                case "System.Collections.Generic.ISet<T>":
                case "System.Collections.Generic.IReadOnlySet<T>":
                    element = argument;
                    shape = YamlShape.Set;
                    return true;
            }
        }
        INamedTypeSymbol? enumerable = null;
        foreach (INamedTypeSymbol candidate in type.AllInterfaces) {
            if (!candidate.IsGenericType || Definition(candidate) != "System.Collections.Generic.IEnumerable<T>") continue;
            if (enumerable is not null) return true;
            enumerable = candidate;
        }
        if (enumerable is null) return false;
        element = enumerable.TypeArguments[0];
        if (!type.IsAbstract && type.TypeKind != TypeKind.Interface && FindInterface(type, "System.Collections.Generic.ICollection<T>") is INamedTypeSymbol collection
            && SymbolEqualityComparer.Default.Equals(collection.TypeArguments[0], element) && HasPublicParameterlessConstructor(type)) {
            shape = YamlShape.CustomCollection;
        }
        return true;
    }

    // ------------------------------------------------------------------ expansion

    private void Expand(YamlTypeEntry entry) {
        Location? location = entry.Type.Locations.FirstOrDefault(static l => l.IsInSource);
        switch (entry.Shape) {
            case YamlShape.Nullable:
            case YamlShape.Array:
            case YamlShape.List:
            case YamlShape.Set:
            case YamlShape.CustomCollection:
                Add(entry.Element!, location);
                break;
            case YamlShape.Dictionary:
            case YamlShape.CustomDictionary:
                Add(entry.Key!, location);
                Add(entry.Value!, location);
                break;
            case YamlShape.Object:
                entry.Object = DescribeObject((INamedTypeSymbol)entry.Type, location);
                if (entry.Object is null) break;
                foreach (YamlPropertyShape property in entry.Object.Properties) {
                    Add(property.Symbol.Type, property.Symbol.Locations.FirstOrDefault() ?? location);
                }
                foreach ((ITypeSymbol derived, _) in entry.Object.Derived) Add(derived, location);
                break;
        }
    }

    private YamlObjectShape? DescribeObject(INamedTypeSymbol type, Location? location) {
        var shape = new YamlObjectShape();
        int errors = _diagnostics.Count;

        foreach (YamlPropertyShape property in GetProperties(type)) shape.Properties.Add(property);
        var explicitNames = new HashSet<string>(StringComparer.Ordinal);
        foreach (YamlPropertyShape property in shape.Properties) {
            if (property.YamlName is not null && !explicitNames.Add(property.YamlName)) Report(YamlDiagnostics.DuplicateName, location, Display(type), property.YamlName);
            if (property.IsRequired) shape.HasRequiredMembers = true;
        }
        shape.HasRequiredMembers |= HasRequiredMembers(type);

        bool isGeneric = IsGenericOrNestedInGeneric(type);
        foreach (YamlPropertyShape property in shape.Properties) {
            IMethodSymbol? setter = property.Symbol.SetMethod;
            if (setter is null || setter.DeclaredAccessibility != Accessibility.Public) {
                property.Setter = YamlSetter.None;
            } else if (!setter.IsInitOnly) {
                property.Setter = YamlSetter.Direct;
            } else if (_unsafeAccessor && (!IsGenericOrNestedInGeneric(setter.ContainingType) || _genericUnsafeAccessor)) {
                property.Setter = YamlSetter.Accessor;
            } else if (!_unsafeAccessor) {
                property.Setter = YamlSetter.Reflection;
            } else {
                Report(YamlDiagnostics.UnsupportedType, location, Display(type),
                    "the init-only property '" + property.Symbol.Name + "' of a generic type can be set only on .NET 9 or later; use a constructor parameter or a settable property");
            }
        }

        if (!type.IsAbstract && type.TypeKind != TypeKind.Interface) SelectConstructor(type, shape, location);

        // Derived types.
        foreach (AttributeData attribute in GetAttributes(type, "OfficeIMO.Yaml.YamlDerivedTypeAttribute")) {
            if (attribute.ConstructorArguments.Length != 2 || attribute.ConstructorArguments[0].Value is not ITypeSymbol derived
                || attribute.ConstructorArguments[1].Value is not string discriminator) continue;
            Location? at = attribute.ApplicationSyntaxReference?.GetSyntax().GetLocation() ?? location;
            if (SymbolEqualityComparer.Default.Equals(derived, type) || !IsAssignable(derived, type)) {
                Report(YamlDiagnostics.InvalidDerivedType, at, Display(type), "'" + Display(derived) + "' does not derive from it");
                continue;
            }
            if (derived.IsAbstract || derived.TypeKind == TypeKind.Interface) {
                Report(YamlDiagnostics.InvalidDerivedType, at, Display(type), "'" + Display(derived) + "' is abstract");
                continue;
            }
            foreach ((ITypeSymbol existing, string existingDiscriminator) in shape.Derived) {
                if (existingDiscriminator == discriminator || SymbolEqualityComparer.Default.Equals(existing, derived)) {
                    Report(YamlDiagnostics.InvalidDerivedType, at, Display(type), "the discriminator '" + discriminator + "' or the type '" + Display(derived) + "' is registered more than once");
                }
            }
            shape.Derived.Add((derived, discriminator));
        }
        AttributeData? polymorphic = GetAttributes(type, "OfficeIMO.Yaml.YamlPolymorphicAttribute").FirstOrDefault();
        if (polymorphic is not null) {
            foreach (KeyValuePair<string, TypedConstant> argument in polymorphic.NamedArguments) {
                if (argument.Key == "DiscriminatorPropertyName") shape.DiscriminatorPropertyName = argument.Value.Value as string;
            }
        }
        return _diagnostics.Count == errors ? shape : null;
    }

    private void SelectConstructor(INamedTypeSymbol type, YamlObjectShape shape, Location? location) {
        IMethodSymbol[] constructors = type.InstanceConstructors.Where(static c => c.DeclaredAccessibility == Accessibility.Public).ToArray();
        IMethodSymbol? marked = null;
        foreach (IMethodSymbol constructor in constructors) {
            if (!GetAttributes(constructor, "OfficeIMO.Yaml.YamlConstructorAttribute").Any()) continue;
            if (marked is not null) {
                Report(YamlDiagnostics.NoConstructor, location, Display(type), "more than one constructor is marked [YamlConstructor]");
                return;
            }
            marked = constructor;
        }
        IMethodSymbol? selected = marked;
        if (selected is null) {
            if (type.IsValueType) {
                shape.UseDefault = true;
                return;
            }
            selected = constructors.FirstOrDefault(static c => c.Parameters.Length == 0);
            if (selected is null && constructors.Length == 1) selected = constructors[0];
        }
        if (selected is null) {
            Report(YamlDiagnostics.NoConstructor, location, Display(type), "it needs a public parameterless constructor, a single public constructor or one marked [YamlConstructor]");
            return;
        }
        if (selected.Parameters.Length == 0) {
            shape.UseDefault = true;
            return;
        }
        foreach (IParameterSymbol parameter in selected.Parameters) {
            bool matched = shape.Properties.Any(p => string.Equals(p.Symbol.Name, parameter.Name, StringComparison.OrdinalIgnoreCase)
                && SymbolEqualityComparer.Default.Equals(Normalize(p.Symbol.Type), Normalize(parameter.Type)));
            if (!matched) {
                Report(YamlDiagnostics.NoConstructor, location, Display(type), "the constructor parameter '" + parameter.Name + "' has no property with the same name and type");
                return;
            }
            if (parameter.RefKind != RefKind.None) {
                Report(YamlDiagnostics.NoConstructor, location, Display(type), "the constructor parameter '" + parameter.Name + "' is passed by reference");
                return;
            }
        }
        shape.Constructor = selected;
    }

    private static bool IsAssignable(ITypeSymbol derived, ITypeSymbol baseType) {
        for (ITypeSymbol? current = derived; current is not null; current = current.BaseType) {
            if (SymbolEqualityComparer.Default.Equals(current, baseType)) return true;
        }
        foreach (INamedTypeSymbol candidate in derived.AllInterfaces) {
            if (SymbolEqualityComparer.Default.Equals(candidate, baseType)) return true;
        }
        return false;
    }

    private static bool IsGenericOrNestedInGeneric(INamedTypeSymbol type) {
        for (INamedTypeSymbol? current = type; current is not null; current = current.ContainingType) {
            if (current.IsGenericType) return true;
        }
        return false;
    }

    private static bool HasRequiredMembers(INamedTypeSymbol type) {
        for (INamedTypeSymbol? current = type; current is not null; current = current.BaseType) {
            foreach (ISymbol member in current.GetMembers()) {
                if (member is IPropertySymbol { IsRequired: true } or IFieldSymbol { IsRequired: true }) return true;
            }
        }
        return false;
    }

    /// <summary>
    /// The serialized properties, as the runtime's <c>YamlTypeShapes.GetProperties</c> lists them: public instance
    /// properties with a public getter, base types first in declaration order, a redeclared property at its base
    /// position, and <c>[YamlIgnore]</c> removing it.
    /// </summary>
    private static List<YamlPropertyShape> GetProperties(INamedTypeSymbol type) {
        var chain = new List<INamedTypeSymbol>();
        if (type.TypeKind == TypeKind.Interface) {
            chain.Add(type);
        } else {
            for (INamedTypeSymbol? current = type; current is not null && current.SpecialType is not (SpecialType.System_Object or SpecialType.System_ValueType); current = current.BaseType) {
                chain.Insert(0, current);
            }
        }
        var result = new List<IPropertySymbol?>();
        var positions = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (INamedTypeSymbol declaring in chain) {
            foreach (ISymbol member in declaring.GetMembers()) {
                if (member is not IPropertySymbol property || property.IsStatic || property.IsIndexer) continue;
                bool anyPublic = property.GetMethod?.DeclaredAccessibility == Accessibility.Public || property.SetMethod?.DeclaredAccessibility == Accessibility.Public;
                if (!anyPublic || property.ExplicitInterfaceImplementations.Length > 0) continue;
                bool ignored = IgnoreCondition(property) == "Always";
                bool readable = property.GetMethod is { DeclaredAccessibility: Accessibility.Public };
                IPropertySymbol? entry = ignored || !readable ? null : property;
                if (positions.TryGetValue(property.Name, out int position)) {
                    result[position] = entry;
                } else {
                    positions.Add(property.Name, result.Count);
                    result.Add(entry);
                }
            }
        }
        var properties = new List<YamlPropertyShape>();
        foreach (IPropertySymbol? property in result) {
            if (property is null) continue;
            properties.Add(new YamlPropertyShape(property) {
                YamlName = PropertyAttribute(property, "OfficeIMO.Yaml.YamlPropertyNameAttribute")?.ConstructorArguments.FirstOrDefault().Value as string,
                IsRequired = property.IsRequired || PropertyAttribute(property, "OfficeIMO.Yaml.YamlRequiredAttribute") is not null,
                IgnoreCondition = IgnoreCondition(property)
            });
        }
        return properties;
    }

    /// <summary>Finds an attribute on a property or the properties it overrides, as reflection's inherited lookup does.</summary>
    private static AttributeData? PropertyAttribute(IPropertySymbol property, string name) {
        for (IPropertySymbol? current = property; current is not null; current = current.OverriddenProperty) {
            AttributeData? found = GetAttributes(current, name).FirstOrDefault();
            if (found is not null) return found;
        }
        return null;
    }

    private static string? IgnoreCondition(IPropertySymbol property) {
        AttributeData? attribute = PropertyAttribute(property, "OfficeIMO.Yaml.YamlIgnoreAttribute");
        if (attribute is null) return null;
        foreach (KeyValuePair<string, TypedConstant> argument in attribute.NamedArguments) {
            if (argument.Key == "Condition" && argument.Value.Value is byte value) return ConditionName(value);
            if (argument.Key == "Condition" && argument.Value.Value is not null) return ConditionName(Convert.ToByte(argument.Value.Value, CultureInfo.InvariantCulture));
        }
        return "Always";
    }

    private static string ConditionName(byte value) => value switch {
        0 => "Never",
        1 => "Always",
        2 => "WhenWritingNull",
        _ => "WhenWritingDefault"
    };

    private static IEnumerable<AttributeData> GetAttributes(ISymbol symbol, string name) =>
        symbol.GetAttributes().Where(a => a.AttributeClass?.ToDisplayString() == name);
}
