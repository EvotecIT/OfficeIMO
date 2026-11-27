using Microsoft.CodeAnalysis;

namespace OfficeIMO.Build.ApiMap.Fluent;

/// <summary>One exported member (one overload) and its proposed 4.0 form.</summary>
internal sealed record LedgerRow(
    string Id,
    string? Package,
    string Receiver,
    string? DeclaredIn,
    string Kind,
    string Member,
    string Signature,
    string? Returns,
    string Treatment,
    string? Proposed,
    string? Semantics,
    string? Reason,
    IReadOnlyList<string>? Flags,
    IReadOnlyList<string>? Frameworks,
    string? Source);

/// <summary>
/// Classifies every exported member of the authoring receivers in one framework build. The rules are the
/// #3344 language and migration maps; lists that are data (families, verbs, retirements, open decisions)
/// come from <see cref="LedgerConfig"/>.
/// </summary>
internal sealed class Ledger {
    internal static class Treatments {
        internal const string Rename = "rename";
        internal const string Generate = "generate";
        internal const string GenerateNested = "generate-nested";
        internal const string GenerateConfigure = "generate-configure";
        internal const string KeepAction = "keep-action";
        internal const string KeepCopy = "keep-copy";
        internal const string KeepAdd = "keep-add";
        internal const string KeepLifecycle = "keep-lifecycle";
        internal const string KeepConversion = "keep-conversion";
        internal const string KeepSelector = "keep-selector";
        internal const string KeepProperty = "keep-property";
        internal const string Keep = "keep";
        internal const string KeepFluent = "keep-fluent";
        internal const string Retire = "retire";
        internal const string Decision = "decision";
    }

    private static readonly SymbolDisplayFormat Signature = new(
        globalNamespaceStyle: SymbolDisplayGlobalNamespaceStyle.Omitted,
        typeQualificationStyle: SymbolDisplayTypeQualificationStyle.NameOnly,
        genericsOptions: SymbolDisplayGenericsOptions.IncludeTypeParameters | SymbolDisplayGenericsOptions.IncludeTypeConstraints,
        memberOptions: SymbolDisplayMemberOptions.IncludeParameters | SymbolDisplayMemberOptions.IncludeModifiers,
        parameterOptions: SymbolDisplayParameterOptions.IncludeType | SymbolDisplayParameterOptions.IncludeName
            | SymbolDisplayParameterOptions.IncludeDefaultValue | SymbolDisplayParameterOptions.IncludeParamsRefOut
            | SymbolDisplayParameterOptions.IncludeExtensionThis,
        propertyStyle: SymbolDisplayPropertyStyle.ShowReadWriteDescriptor,
        miscellaneousOptions: SymbolDisplayMiscellaneousOptions.UseSpecialTypes | SymbolDisplayMiscellaneousOptions.IncludeNullableReferenceTypeModifier);

    private static readonly SymbolDisplayFormat TypeName = new(
        globalNamespaceStyle: SymbolDisplayGlobalNamespaceStyle.Omitted,
        typeQualificationStyle: SymbolDisplayTypeQualificationStyle.NameAndContainingTypes,
        genericsOptions: SymbolDisplayGenericsOptions.IncludeTypeParameters,
        miscellaneousOptions: SymbolDisplayMiscellaneousOptions.UseSpecialTypes | SymbolDisplayMiscellaneousOptions.IncludeNullableReferenceTypeModifier);

    private readonly LedgerConfig _config;
    private readonly SourceLocationIndex _source;
    private readonly Dictionary<INamedTypeSymbol, List<IMethodSymbol>> _extensions = new(SymbolEqualityComparer.Default);
    private readonly Dictionary<INamedTypeSymbol, List<INamedTypeSymbol>> _derived = new(SymbolEqualityComparer.Default);

    internal int PublicTypes { get; private set; }
    internal int Receivers { get; private set; }

    internal Ledger(LedgerConfig config, SourceLocationIndex source) {
        _config = config;
        _source = source;
    }

    internal IReadOnlyList<LedgerRow> Build(MetadataCatalog catalog, string framework) {
        var types = catalog.Types.Where(type => type.TypeKind is TypeKind.Class or TypeKind.Struct or TypeKind.Interface).ToArray();
        PublicTypes = catalog.Types.Count;
        foreach (var type in types) {
            if (type.IsStatic) {
                foreach (var method in type.GetMembers().OfType<IMethodSymbol>().Where(method => method.IsExtensionMethod && IsPublic(method))) {
                    if (ReceiverOf(method) is not { } receiver) continue;
                    if (!_extensions.TryGetValue(receiver, out var list)) _extensions[receiver] = list = [];
                    list.Add(method);
                }
            }
            for (var parent = type.BaseType?.OriginalDefinition; parent != null; parent = parent.BaseType?.OriginalDefinition) {
                if (!_derived.TryGetValue(parent, out var list)) _derived[parent] = list = [];
                list.Add(type);
            }
        }
        var rows = new List<LedgerRow>();
        var receivers = new HashSet<INamedTypeSymbol>(SymbolEqualityComparer.Default);
        foreach (var type in types.Where(IsReceiver)) {
            receivers.Add(type);
            foreach (var member in DeclaredMembers(type)) rows.Add(Classify(type, member, framework));
        }
        foreach (var (receiver, methods) in _extensions) {
            if (!receivers.Contains(receiver) && !methods.Any(method => IsAuthoringName(method.Name) || ReturnsReceiver(method, receiver))) continue;
            receivers.Add(receiver);
            foreach (var method in methods) rows.Add(Classify(receiver, method, framework));
        }
        // Static entry points carry decisions too (the conversion request), without treating every helper class as a receiver.
        foreach (var type in types.Where(type => type.IsStatic && _config.Decisions.Concat<object>(_config.Retire).Any(rule => RuleTargets(rule, type)))) {
            receivers.Add(type);
            foreach (var method in DeclaredMembers(type).OfType<IMethodSymbol>().Where(method => !method.IsExtensionMethod)) rows.Add(Classify(type, method, framework));
        }
        Receivers = receivers.Count;
        return rows;
    }

    private bool RuleTargets(object rule, INamedTypeSymbol type) => rule switch {
        DecisionRule decision => LedgerConfig.ReceiverMatches(decision.Receiver, Name(type)),
        Retirement retirement => LedgerConfig.ReceiverMatches(retirement.Receiver, Name(type)),
        _ => false
    };

    private bool IsReceiver(INamedTypeSymbol type) {
        if (type.IsStatic || Inherits(type, "System.Exception") || Inherits(type, "System.Attribute")) return false;
        foreach (var member in type.GetMembers()) {
            if (!IsPublic(member) || member.IsStatic) continue;
            if (member is IPropertySymbol property && IsPublicSetter(property)) return true;
            if (member is IMethodSymbol { MethodKind: MethodKind.Ordinary } method && (IsAuthoringName(method.Name) || ReturnsReceiver(method, type))) return true;
        }
        return false;
    }

    private bool IsAuthoringName(string name) =>
        LedgerConfig.MatchPrefix(_config.SetPrefixes, name) != null || LedgerConfig.MatchPrefix(_config.CopyPrefixes, name) != null
        || LedgerConfig.MatchPrefix(_config.AddPrefixes, name) != null;

    private static IEnumerable<ISymbol> DeclaredMembers(INamedTypeSymbol type) => type.GetMembers().Where(member =>
        IsPublic(member) && !member.IsImplicitlyDeclared && member is not INamedTypeSymbol
        && member is not IMethodSymbol { MethodKind: MethodKind.PropertyGet or MethodKind.PropertySet or MethodKind.EventAdd or MethodKind.EventRemove or MethodKind.StaticConstructor or MethodKind.Destructor });

    private LedgerRow Classify(INamedTypeSymbol receiver, ISymbol member, string framework) => member switch {
        IMethodSymbol method => ClassifyMethod(receiver, method, framework),
        IPropertySymbol property => ClassifyProperty(receiver, property, framework),
        IFieldSymbol field => Row(receiver, field, framework, field.IsConst || field.IsReadOnly || field.IsStatic ? Treatments.Keep : Treatments.Decision, null, "",
            field.IsConst || field.IsReadOnly || field.IsStatic ? null : "Public mutable field: make it a property or document the exception.",
            field.IsConst || field.IsReadOnly || field.IsStatic ? [] : ["public-field"]),
        IEventSymbol @event => Row(receiver, @event, framework, Treatments.Keep, null, "", null, []),
        _ => Row(receiver, member, framework, Treatments.Keep, null, "", null, [])
    };

    private LedgerRow ClassifyMethod(INamedTypeSymbol receiver, IMethodSymbol method, string framework) {
        string name = method.Name;
        string receiverName = Name(receiver);
        string semantics = method.MethodKind == MethodKind.Ordinary ? BodyFacts.Semantics(_source.FindSyntax(method), method) : "";
        var flags = new List<string>();
        if (receiver.TypeKind == TypeKind.Struct) flags.Add("struct-receiver");
        if (receiver.TypeKind == TypeKind.Interface && !method.IsExtensionMethod) flags.Add("interface-member");

        if (method.MethodKind is MethodKind.UserDefinedOperator or MethodKind.Conversion || method.IsOverride || _config.ObjectContract.Contains(name))
            return Row(receiver, method, framework, Treatments.Keep, null, semantics, null, flags);
        if (method.MethodKind == MethodKind.Constructor)
            return Row(receiver, method, framework, Treatments.KeepLifecycle, null, semantics, null, flags);
        if (name is "Dispose" or "DisposeAsync")
            return Row(receiver, method, framework, Treatments.KeepLifecycle, null, semantics, null, flags);
        if (_config.Retire.FirstOrDefault(rule => rule.Member == name && LedgerConfig.ReceiverMatches(rule.Receiver, receiverName)) is { } retirement)
            return Row(receiver, method, framework, Treatments.Retire, retirement.Replacement, semantics, retirement.Reason, flags);
        if (_config.Decisions.FirstOrDefault(rule => rule.Members.Contains(name) && LedgerConfig.ReceiverMatches(rule.Receiver, receiverName)) is { } decision) {
            flags.Add(decision.Code);
            return Row(receiver, method, framework, Treatments.Decision, decision.Proposal, semantics, decision.Reason, flags);
        }
        if (LedgerConfig.MatchPrefix(_config.SetPrefixes, name) is { } set) {
            if (name.Length > set.Length) return ClassifySet(receiver, method, framework, name[set.Length..], semantics, flags);
            flags.Add("bare-set");
            return Row(receiver, method, framework, Treatments.Decision, null, semantics, "A bare Set(...) has no property name to take: name the configured value or keep it as an action.", flags);
        }
        if (LedgerConfig.MatchPrefix(_config.CopyPrefixes, name) is { } with)
            return ClassifyWith(receiver, method, framework, name[with.Length..], semantics, flags);
        if (LedgerConfig.MatchPrefix(["Use"], name) is { } use && name.Length > use.Length) {
            flags.Add("use-prefix");
            var parameters = method.Parameters.Skip(method.IsExtensionMethod ? 1 : 0).ToArray();
            return Row(receiver, method, framework, Treatments.Decision, ProposedCall(name[use.Length..], parameters), semantics,
                "Use... configuration has no #3344 family: drop the prefix like Set..., or keep Use for selecting a profile or strategy.", flags);
        }
        if (LedgerConfig.MatchPrefix(_config.AddPrefixes, name) != null) {
            if (ReturnsReceiver(method, receiver)) flags.Add("returns-parent");
            else if (method.ReturnsVoid) flags.Add("returns-void");
            return Row(receiver, method, framework, Treatments.KeepAdd, null, semantics,
                ReturnsReceiver(method, receiver) ? "Collection or parent-continuing addition; keep, or move to a Content scope." : null, flags);
        }
        if (_config.LifecycleNames.Contains(name))
            return Row(receiver, method, framework, Treatments.KeepLifecycle, null, semantics, null, flags);
        if (LedgerConfig.MatchPrefix(_config.ConversionPrefixes, name) != null)
            return Row(receiver, method, framework, Treatments.KeepConversion, null, semantics, null, flags);
        if (LedgerConfig.MatchPrefix(_config.SelectorPrefixes, name) != null)
            return Row(receiver, method, framework, Treatments.KeepSelector, null, semantics, null, flags);
        if (LedgerConfig.MatchPrefix(_config.ActionPrefixes, name) != null)
            return Row(receiver, method, framework, Treatments.KeepAction, null, semantics, null, flags);
        if (method.IsStatic && !method.IsExtensionMethod && ReturnsReceiver(method, receiver))
            return Row(receiver, method, framework, Treatments.KeepLifecycle, null, semantics, "Static factory.", flags);
        if (ReturnsReceiver(method, receiver)) {
            if (FindProperty(receiver, name) != null) flags.Add("matches-property");
            if (method.Parameters.Any(parameter => parameter.Type.TypeKind == TypeKind.Delegate)) flags.Add("callback");
            return Row(receiver, method, framework, Treatments.KeepFluent, null, semantics, null, flags);
        }
        return Row(receiver, method, framework, Treatments.Keep, null, semantics, null, flags);
    }

    private LedgerRow ClassifySet(INamedTypeSymbol receiver, IMethodSymbol method, string framework, string suffix, string semantics, List<string> flags) {
        string package = method.ContainingAssembly.Name;
        var family = _config.Families.FirstOrDefault(rule => rule.Old == method.Name && (rule.Package == null || rule.Package == package));
        string target = family?.New ?? suffix;
        string treatment = family?.Treatment ?? Treatments.Rename;
        string? reason = family?.Reason;
        var parameters = method.Parameters.Skip(method.IsExtensionMethod ? 1 : 0).ToArray();
        var property = FindProperty(receiver, suffix) ?? FindProperty(receiver, "Is" + suffix);

        if (method.IsGenericMethod) {
            flags.Add("generic-set");
            treatment = Treatments.Decision;
            reason = "Generic Set method: the extension needs its type parameters and constraints, which a generated mutator call does not carry; keep it handwritten or extend the generator.";
        } else if (method.IsStatic && !method.IsExtensionMethod) {
            flags.Add("static-set");
            treatment = Treatments.Decision;
            reason = "Static Set method: a shared default or factory, not receiver configuration.";
        } else if (!method.ReturnsVoid && !ReturnsReceiver(method, receiver)) {
            flags.Add("returns-other");
            treatment = Treatments.Decision;
            reason = $"Returns {Display(method.ReturnType)}, not the receiver: a property-named extension would change the result.";
        } else if (property == null && parameters.Length > 1) {
            if (IsKey(parameters[0])) {
                flags.Add("keyed-set");
                treatment = Treatments.KeepAction;
                reason = $"Writes a value addressed by '{parameters[0].Name}': a keyed action keeps its Set verb.";
            } else {
                flags.Add("multi-value");
            }
        }
        if (method.ReturnsVoid) flags.Add("void-setter");
        if (property == null) flags.Add("no-property");
        if (semantics == BodyFacts.ReturnsNew) {
            flags.Add("set-returns-copy");
            treatment = Treatments.Decision;
            reason = "Set method returns a new instance: copy semantics under a mutation name.";
        }
        if (treatment != Treatments.Decision && Collision(receiver, target, method) is { } collision) {
            flags.Add("name-collision");
            treatment = Treatments.Decision;
            reason = $"{target} already names {collision} on the receiver or a derived receiver.";
        }
        return Row(receiver, method, framework, treatment, ProposedCall(target, parameters), semantics, reason, flags);
    }

    private LedgerRow ClassifyWith(INamedTypeSymbol receiver, IMethodSymbol method, string framework, string suffix, string semantics, List<string> flags) {
        bool immutable = IsImmutable(receiver);
        if (semantics is BodyFacts.ReturnsReceiver or BodyFacts.Mixed) {
            flags.Add("with-mutates");
            var parameters = method.Parameters.Skip(method.IsExtensionMethod ? 1 : 0).ToArray();
            return Row(receiver, method, framework, Treatments.Decision, ProposedCall(suffix, parameters), semantics,
                "With... returns the same mutated receiver: rename to property-named configuration, or make it a real copy.", flags);
        }
        if (!immutable && semantics != BodyFacts.ReturnsNew) flags.Add("copy-unverified");
        return Row(receiver, method, framework, Treatments.KeepCopy, null, semantics, null, flags);
    }

    private LedgerRow ClassifyProperty(INamedTypeSymbol receiver, IPropertySymbol property, string framework) {
        var flags = new List<string>();
        if (receiver.TypeKind == TypeKind.Struct) flags.Add("struct-receiver");
        if (property.IsIndexer) return Row(receiver, property, framework, Treatments.KeepSelector, null, "", null, flags);
        if (property.IsStatic) return Row(receiver, property, framework, Treatments.Keep, null, "", null, flags);

        string receiverName = Name(receiver);
        string package = property.ContainingAssembly.Name;
        var rename = _config.PropertyRenames.FirstOrDefault(rule => rule.Old == property.Name
            && (rule.Package == null || rule.Package == package) && LedgerConfig.ReceiverMatches(rule.Receiver, receiverName));
        string name = rename?.New ?? property.Name;
        if (rename != null) flags.Add("property-rename");

        if (IsPublicSetter(property)) {
            if (property.Type.TypeKind == TypeKind.Delegate)
                return Row(receiver, property, framework, Treatments.GenerateConfigure, $"Configure(x => x.{name} = value)", "",
                    "Delegate-valued property is assigned through Configure, never invoked by a same-named call.", flags);
            bool boolean = IsBoolean(property.Type);
            string target = name;
            if (boolean && name.Length > 2 && name.StartsWith("Is", StringComparison.Ordinal) && char.IsUpper(name[2])) {
                target = name[2..];
                flags.Add("is-prefix");
            }
            if (receiver.TypeKind == TypeKind.Struct) {
                flags.Add("struct-property");
                return Row(receiver, property, framework, Treatments.Decision, null, "",
                    "Setter on a value type: an extension would write to a copy. Route through a handle mutator or leave as assignment.", flags);
            }
            string? reason = rename?.Reason;
            if (HasMethod(receiver, "Set" + property.Name) || HasMethod(receiver, "Set" + target)) flags.Add("merges-set");
            string treatment = Treatments.Generate;
            if (Collision(receiver, target, property) is { } collision) {
                flags.Add("name-collision");
                treatment = Treatments.Decision;
                reason = $"{target} already names {collision} on the receiver or a derived receiver.";
            }
            string proposed = boolean ? $"{target}({Display(property.Type)} value = true)" : $"{target}({Display(property.Type)} value)";
            return Row(receiver, property, framework, treatment, proposed, "", reason, flags);
        }
        if (property.SetMethod is { IsInitOnly: true }) {
            flags.Add("init-only");
            return Row(receiver, property, framework, Treatments.KeepProperty, null, "", null, flags);
        }
        if (property.SetMethod != null) {
            flags.Add(HasMethod(receiver, "Set" + property.Name) ? "set-method" : "non-public-setter");
            return Row(receiver, property, framework, Treatments.KeepProperty, null, "",
                HasMethod(receiver, "Set" + property.Name) ? "Configured through the Set method's rename." : "Setter is not public: read-only state unless a mutator is exposed.", flags);
        }
        if (rename != null) {
            string? clash = Collision(receiver, name, property, includeProperties: true);
            if (clash != null) flags.Add("name-collision");
            return Row(receiver, property, framework, clash == null ? Treatments.Rename : Treatments.Decision, name, "",
                clash == null ? rename.Reason : $"{rename.Reason} {name} already names {clash}.", flags);
        }
        if (IsOptionObject(property.Type) && BodyFacts.GetterReturnsCopy(_source.FindSyntax(property))) {
            flags.Add("getter-copies");
            return Row(receiver, property, framework, Treatments.KeepSelector, null, "",
                "The getter returns a copy, so an in-place callback would change nothing; change it through the receiver's own With... or Set... method.", flags);
        }
        if (IsOptionObject(property.Type))
            return Row(receiver, property, framework, Treatments.GenerateNested, $"{name}({Display(property.Type).TrimEnd('?')} => ...)", "",
                "Read-only option object configured in place through a nested callback.", flags);
        if (IsCollection(property.Type)) flags.Add("collection");
        return Row(receiver, property, framework, Treatments.KeepSelector, null, "", null, flags);
    }

    /// <summary>Names an existing member that a generated extension called <paramref name="target"/> could not coexist with.</summary>
    private string? Collision(INamedTypeSymbol receiver, string target, ISymbol source, bool includeProperties = false) {
        foreach (var type in SelfAndBases(receiver).Concat(Derived(receiver))) {
            foreach (var member in type.GetMembers(target)) {
                if (SymbolEqualityComparer.Default.Equals(member, source)) continue;
                // A same-named property or field does not hide an extension method (List<T>.Count / Count()).
                if (member is IPropertySymbol or IFieldSymbol && !includeProperties) continue;
                if (member is IMethodSymbol existing) {
                    if (!existing.IsStatic && ReturnsReceiver(existing, type)) continue;
                    return (existing.IsStatic ? "static method " : "method ") + Name(type) + "." + existing.Name;
                }
                if (member is INamedTypeSymbol nested) return "nested type " + Name(nested);
                if (member is IEventSymbol) return "event " + Name(type) + "." + member.Name;
            }
        }
        return null;
    }

    private IEnumerable<INamedTypeSymbol> Derived(INamedTypeSymbol type) =>
        _derived.TryGetValue(type.OriginalDefinition, out var list) ? list : [];

    private static IEnumerable<INamedTypeSymbol> SelfAndBases(INamedTypeSymbol type) {
        for (var current = type; current != null && current.SpecialType != SpecialType.System_Object; current = current.BaseType) yield return current;
        foreach (var contract in type.AllInterfaces) yield return contract;
    }

    private bool HasMethod(INamedTypeSymbol receiver, string name) =>
        SelfAndBases(receiver).Any(type => type.GetMembers(name).OfType<IMethodSymbol>().Any(IsPublic))
        || _extensions.TryGetValue(receiver.OriginalDefinition, out var list) && list.Any(method => method.Name == name);

    private static IPropertySymbol? FindProperty(INamedTypeSymbol receiver, string name) =>
        SelfAndBases(receiver).SelectMany(type => type.GetMembers(name)).OfType<IPropertySymbol>().FirstOrDefault(property => IsPublic(property) && !property.IsStatic);

    private static bool ReturnsReceiver(IMethodSymbol method, INamedTypeSymbol receiver) {
        var returned = method.ReturnType;
        if (returned is ITypeParameterSymbol parameter && method.IsExtensionMethod && SymbolEqualityComparer.Default.Equals(method.Parameters[0].Type, parameter)) return true;
        return returned is INamedTypeSymbol named && SymbolEqualityComparer.Default.Equals(named.OriginalDefinition, receiver.OriginalDefinition);
    }

    private static INamedTypeSymbol? ReceiverOf(IMethodSymbol method) {
        var type = method.Parameters[0].Type;
        if (type is ITypeParameterSymbol parameter) type = parameter.ConstraintTypes.FirstOrDefault() ?? type;
        return (type as INamedTypeSymbol)?.OriginalDefinition;
    }

    private static bool IsImmutable(INamedTypeSymbol type) =>
        type.IsRecord && !type.GetMembers().OfType<IPropertySymbol>().Any(IsPublicSetter) || type.IsReadOnly
        || !type.GetMembers().OfType<IPropertySymbol>().Any(IsPublicSetter) && !type.GetMembers().OfType<IFieldSymbol>().Any(field => IsPublic(field) && !field.IsReadOnly && !field.IsConst);

    private static bool IsOptionObject(ITypeSymbol type) =>
        type is INamedTypeSymbol { TypeKind: TypeKind.Class, SpecialType: SpecialType.None } named
        && !named.ContainingNamespace.ToDisplayString().StartsWith("System", StringComparison.Ordinal)
        && !IsCollection(named) && named.GetMembers().OfType<IPropertySymbol>().Any(IsPublicSetter);

    private static bool IsCollection(ITypeSymbol type) =>
        type.SpecialType != SpecialType.System_String && type.AllInterfaces.Any(contract => contract.OriginalDefinition.ToDisplayString() == "System.Collections.IEnumerable" || contract.ToDisplayString() == "System.Collections.IEnumerable");

    private static bool IsBoolean(ITypeSymbol type) =>
        type.SpecialType == SpecialType.System_Boolean
        || type is INamedTypeSymbol { OriginalDefinition.SpecialType: SpecialType.System_Nullable_T } nullable && nullable.TypeArguments[0].SpecialType == SpecialType.System_Boolean;

    private static bool IsPublicSetter(IPropertySymbol property) =>
        IsPublic(property) && !property.IsStatic && property.SetMethod is { DeclaredAccessibility: Accessibility.Public, IsInitOnly: false };

    private static bool IsPublic(ISymbol symbol) => symbol.DeclaredAccessibility == Accessibility.Public;

    private static bool Inherits(INamedTypeSymbol type, string name) {
        for (var current = type.BaseType; current != null; current = current.BaseType) if (current.ToDisplayString() == name) return true;
        return false;
    }

    /// <summary>A first parameter that addresses what is written (an index, key, name or coordinate) rather than being the value.</summary>
    private static bool IsKey(IParameterSymbol parameter) {
        string name = parameter.Name.ToLowerInvariant();
        return name is "x" or "row" or "column" or "col" or "level" or "point" or "series" or "slot" or "position" or "key" or "name" or "id" or "tag" or "section" or "cell" or "address" or "reference"
            || name.EndsWith("index", StringComparison.Ordinal) || name.EndsWith("key", StringComparison.Ordinal) || name.EndsWith("name", StringComparison.Ordinal)
            || name.EndsWith("id", StringComparison.Ordinal) || name.EndsWith("number", StringComparison.Ordinal) || name.EndsWith("row", StringComparison.Ordinal)
            || name.EndsWith("column", StringComparison.Ordinal);
    }

    private static string ProposedCall(string name, IReadOnlyList<IParameterSymbol> parameters) {
        if (parameters.Count == 1 && IsBoolean(parameters[0].Type) && !parameters[0].HasExplicitDefaultValue)
            return $"{name}({Display(parameters[0].Type)} {parameters[0].Name} = true)";
        return name + "(" + string.Join(", ", parameters.Select(parameter => parameter.ToDisplayString(Signature))) + ")";
    }

    private static string Display(ITypeSymbol type) => type.ToDisplayString(TypeName);
    private static string Name(INamedTypeSymbol type) => type.OriginalDefinition.ToDisplayString(SymbolDisplayFormat.CSharpErrorMessageFormat.WithGenericsOptions(SymbolDisplayGenericsOptions.None));

    private LedgerRow Row(INamedTypeSymbol receiver, ISymbol member, string framework, string treatment, string? proposed, string semantics, string? reason, IReadOnlyList<string> flags) {
        bool extension = member is IMethodSymbol { IsExtensionMethod: true };
        string kind = member switch {
            IMethodSymbol { MethodKind: MethodKind.Constructor } => "constructor",
            IMethodSymbol { IsExtensionMethod: true } => "extension",
            IMethodSymbol { MethodKind: MethodKind.UserDefinedOperator or MethodKind.Conversion } => "operator",
            IMethodSymbol => member.IsStatic ? "static-method" : "method",
            IPropertySymbol { IsIndexer: true } => "indexer",
            IPropertySymbol => "property",
            IFieldSymbol => "field",
            IEventSymbol => "event",
            _ => member.Kind.ToString().ToLowerInvariant()
        };
        string returns = member switch {
            IMethodSymbol { MethodKind: MethodKind.Constructor } => "",
            IMethodSymbol method => Display(method.ReturnType),
            IPropertySymbol property => Display(property.Type),
            IFieldSymbol field => Display(field.Type),
            IEventSymbol @event => Display(@event.Type),
            _ => ""
        };
        var location = _source.Find(member).FirstOrDefault();
        return new LedgerRow(
            member.GetDocumentationCommentId() ?? member.ToDisplayString(),
            member.ContainingAssembly.Name,
            Name(receiver),
            extension ? Name(member.ContainingType) : null,
            kind,
            member.Name,
            member.ToDisplayString(Signature),
            returns.Length == 0 ? null : returns,
            treatment,
            proposed,
            semantics.Length == 0 ? null : semantics,
            reason,
            flags.Count == 0 ? null : flags.Distinct().ToArray(),
            [framework],
            location == null ? null : location.File + ":" + location.Line);
    }
}
