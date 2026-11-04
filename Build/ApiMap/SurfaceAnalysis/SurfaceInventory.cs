using Microsoft.CodeAnalysis;

namespace OfficeIMO.Build.ApiMap;

internal sealed record SurfaceCandidate(string Id, string Kind, IReadOnlyList<SurfaceLocation> Locations, string Reason, int MembersHiddenWithType = 0);
internal sealed record SurfaceReference(string Id, string Source, IReadOnlyList<string> Evidence);
internal sealed record PackageSurface(string Package, int PublicTypes, int PublicMembers, int EvidenceFiles,
    IReadOnlyList<SurfaceCandidate> Candidates, IReadOnlyList<SurfaceReference> RetainedReferences);

/// <summary>Builds one documentation and signature closure across all selected assemblies.</summary>
internal static class SurfaceInventory {
    internal static IReadOnlyList<PackageSurface> Create(string root, string[] files, string[] packages) {
        var catalog = new MetadataCatalog(root, packages);
        var locations = new SourceLocationIndex(root, files, packages.ToHashSet(StringComparer.Ordinal));
        var documentation = new Dictionary<ISymbol, DocumentationIndex>(SymbolEqualityComparer.Default);
        foreach (string package in packages) documentation.Add(catalog.Assemblies[package], new DocumentationIndex(root, files, package));
        var types = catalog.Types.ToDictionary(type => (ISymbol)type.OriginalDefinition, type => type, SymbolEqualityComparer.Default);
        var retainedTypes = new HashSet<ISymbol>(SymbolEqualityComparer.Default);
        var models = new HashSet<ISymbol>(SymbolEqualityComparer.Default);
        var retainedMembers = new HashSet<ISymbol>(SymbolEqualityComparer.Default);
        var reasons = new Dictionary<ISymbol, SurfaceReference>(SymbolEqualityComparer.Default);
        var pending = new Queue<INamedTypeSymbol>();
        var signatureTypes = new HashSet<ISymbol>(SymbolEqualityComparer.Default);

        void KeepType(INamedTypeSymbol type, string reason, IReadOnlyList<string>? evidence = null) {
            type = type.OriginalDefinition;
            if (!types.ContainsKey(type)) return;
            if (retainedTypes.Add(type)) reasons[type] = Reference(type, reason, evidence);
            if (type.ContainingType != null) KeepType(type.ContainingType, "Contains documented nested type " + Id(type));
        }
        void KeepModel(ITypeSymbol type, string source) {
            if (!signatureTypes.Add(type)) return;
            if (type is IArrayTypeSymbol array) { KeepModel(array.ElementType, source); return; }
            if (type is IPointerTypeSymbol pointer) { KeepModel(pointer.PointedAtType, source); return; }
            if (type is ITypeParameterSymbol parameter) {
                foreach (var constraint in parameter.ConstraintTypes) KeepModel(constraint, source);
                return;
            }
            if (type is not INamedTypeSymbol named) return;
            foreach (var argument in named.TypeArguments) KeepModel(argument, source);
            named = named.OriginalDefinition;
            if (!types.ContainsKey(named)) return;
            KeepType(named, "Model exposed by " + source);
            if (models.Add(named)) pending.Enqueue(named);
        }
        void Follow(ISymbol symbol) {
            foreach (var type in SignatureTypes(symbol)) KeepModel(type, Id(symbol));
        }
        void KeepMember(ISymbol member, string reason, IReadOnlyList<string>? evidence = null) {
            KeepType(member.ContainingType, reason, evidence);
            if (!retainedMembers.Add(member)) return;
            reasons[member] = Reference(member, reason, evidence);
            Follow(member);
        }

        foreach (var type in catalog.Types) {
            var docs = documentation[type.ContainingAssembly];
            var typeEvidence = docs.Find(type.Name);
            bool syntaxNode = type.ContainingNamespace.ToDisplayString() == "OfficeIMO.Markdown"
                && type.AllInterfaces.Any(contract => contract.Name is "IMarkdownBlock" or "IMarkdownInline");
            bool externalContract = Inherits(type, "System.Exception") || Inherits(type, "System.Attribute");
            if (syntaxNode || externalContract)
                KeepType(type, syntaxNode ? "Public Markdown syntax tree" : "Caller exception or attribute contract");
            if (typeEvidence.Count > 0) {
                KeepType(type, "Documented type", typeEvidence);
                Follow(type);
            }
            foreach (var member in PublicMembers(type)) {
                if (IsUnifiedDocumentContract(member))
                    KeepMember(member, "Unified document lifecycle or conversion contract");
                bool constructor = member is IMethodSymbol { MethodKind: MethodKind.Constructor };
                var evidence = constructor ? docs.FindConstruction(type.Name) : docs.Find(member.Name);
                bool extension = member is IMethodSymbol { IsExtensionMethod: true };
                if (evidence.Count > 0 && (typeEvidence.Count > 0 || extension))
                    KeepMember(member, constructor ? "Documented construction" : "Documented member name", evidence);
            }
        }
        // Type roots also contribute their public data models. Interfaces and overrides
        // retain their contracts; owning document methods still need user documentation.
        foreach (var type in catalog.Types.Where(type => retainedTypes.Contains(type)).ToArray())
            if (models.Add(type)) pending.Enqueue(type);
        while (pending.Count > 0) {
            var type = pending.Dequeue();
            Follow(type);
            if (type.IsAbstract && type.TypeKind == TypeKind.Class) {
                foreach (var subtype in catalog.Types.Where(candidate =>
                    SymbolEqualityComparer.Default.Equals(candidate.BaseType?.OriginalDefinition, type.OriginalDefinition)))
                    KeepModel(subtype, "Concrete model exposed through " + Id(type));
            }
            var implementations = type.AllInterfaces.SelectMany(contract => contract.GetMembers())
                .Select(type.FindImplementationForInterfaceMember).Where(member => member != null)
                .ToHashSet(SymbolEqualityComparer.Default);
            foreach (var member in PublicMembers(type)) {
                bool data = member is IPropertySymbol or IFieldSymbol or IEventSymbol;
                bool contract = type.TypeKind is TypeKind.Interface or TypeKind.Delegate || member.IsOverride || member.IsVirtual || member.IsAbstract || implementations.Contains(member);
                // Referenced models need their constructors even when their names do not end
                // in Options/Info (for example, theme palettes and font definitions).
                bool modelConstructor = member is IMethodSymbol { MethodKind: MethodKind.Constructor } constructor
                    && !constructor.Parameters.Any(parameter => IsDocumentOwner(parameter.Type));
                bool valueContract = type.IsValueType;
                bool modelTransformation = member is IMethodSymbol { IsStatic: false, MethodKind: MethodKind.Ordinary } method
                    && method.Name.StartsWith("With", StringComparison.Ordinal)
                    && SymbolEqualityComparer.Default.Equals(method.ReturnType.OriginalDefinition, type.OriginalDefinition);
                contract |= member is IMethodSymbol { MethodKind: MethodKind.UserDefinedOperator or MethodKind.Conversion };
                var memberEvidence = member is IMethodSymbol { MethodKind: MethodKind.Constructor }
                    ? documentation[type.ContainingAssembly].FindConstruction(type.Name) : documentation[type.ContainingAssembly].Find(member.Name);
                if (memberEvidence.Count > 0) KeepMember(member, "Documented model member", memberEvidence);
                bool generatedRecord = type.IsRecord && locations.Find(member).Count == 0;
                if (data || contract || modelConstructor || valueContract || modelTransformation || generatedRecord)
                    KeepMember(member, contract ? "Public type contract" : "Exposed model data");
            }
        }
        return packages.Select(package => {
            var assembly = catalog.Assemblies[package];
            var packageTypes = catalog.Types.Where(type => SymbolEqualityComparer.Default.Equals(type.ContainingAssembly, assembly)).ToArray();
            var candidates = new List<SurfaceCandidate>();
            foreach (var type in packageTypes) {
                var members = PublicMembers(type).ToArray();
                if (!retainedTypes.Contains(type)) {
                    candidates.Add(new SurfaceCandidate(Id(type), "type", locations.Find(type),
                        "No documentation reference or exposed model signature", members.Length));
                    continue;
                }
                foreach (var member in members.Where(member => !retainedMembers.Contains(member)))
                    candidates.Add(new SurfaceCandidate(Id(member), member.Kind.ToString(), locations.Find(member),
                        "No member reference in documentation or exposed model contract"));
            }
            return new PackageSurface(package, packageTypes.Length, packageTypes.Sum(type => PublicMembers(type).Count()),
                documentation[assembly].FileCount, candidates.OrderBy(candidate => candidate.Id, StringComparer.Ordinal).ToArray(),
                reasons.Where(item => SymbolEqualityComparer.Default.Equals(item.Key.ContainingAssembly, assembly))
                    .Select(item => item.Value).OrderBy(reference => reference.Id, StringComparer.Ordinal).ToArray());
        }).ToArray();
    }

    private static bool Inherits(INamedTypeSymbol type, string baseName) {
        for (var current = type.BaseType; current != null; current = current.BaseType)
            if (current.ToDisplayString() == baseName) return true;
        return false;
    }

    private static bool IsDocumentOwner(ITypeSymbol type) => type.ToDisplayString() is
        "OfficeIMO.Word.WordDocument" or "OfficeIMO.Excel.ExcelDocument" or
        "OfficeIMO.PowerPoint.PowerPointDocument" or "OfficeIMO.Pdf.PdfDocument" or
        "OfficeIMO.Markdown.MarkdownDocument";

    // Required lifecycle and conversion overloads remain public even when an example
    // demonstrates only the synchronous form. Static file-converter helpers do not
    // acquire a public contract merely because their names resemble these methods.
    private static bool IsUnifiedDocumentContract(ISymbol symbol) {
        if (symbol is not IMethodSymbol method) return false;
        string name = method.Name.EndsWith("Async", StringComparison.Ordinal)
            ? method.Name[..^5] : method.Name;
        if (IsDocumentOwner(method.ContainingType)
            && name is "Create" or "Load" or "Save" or "Inspect") return true;
        if (method.ContainingType.ToDisplayString() == "OfficeIMO.OfficeConverter"
            && name is "Convert" or "ConvertDirectory" or "Analyze") return true;
        bool instanceDocument = !method.IsStatic && IsDocumentOwner(method.ContainingType);
        bool domainExtension = method.IsExtensionMethod && method.Parameters.Length > 0
            && method.Parameters[0].Type.ContainingNamespace.ToDisplayString()
                .StartsWith("OfficeIMO", StringComparison.Ordinal);
        if (!instanceDocument && !domainExtension) return false;
        string? target = name.StartsWith("SaveAs", StringComparison.Ordinal) ? name[6..]
            : name.StartsWith("To", StringComparison.Ordinal) ? name[2..] : null;
        return target is "Word" or "Excel" or "PowerPoint" or "Pdf" or "Html" or
            "Markdown" or "Image" or "Images";
    }

    private static string Id(ISymbol symbol) => symbol.GetDocumentationCommentId() ?? symbol.ToDisplayString();
    private static SurfaceReference Reference(ISymbol symbol, string source, IReadOnlyList<string>? evidence) => new(Id(symbol), source, evidence ?? []);
    private static IEnumerable<ISymbol> PublicMembers(INamedTypeSymbol type) => type.GetMembers().Where(member =>
        member.DeclaredAccessibility is Accessibility.Public or Accessibility.Protected or Accessibility.ProtectedOrInternal
        && member is not INamedTypeSymbol && !member.IsImplicitlyDeclared
        && member is not IMethodSymbol { MethodKind: MethodKind.PropertyGet or MethodKind.PropertySet or MethodKind.EventAdd or MethodKind.EventRemove or MethodKind.StaticConstructor });

    private static IEnumerable<ITypeSymbol> SignatureTypes(ISymbol symbol) {
        switch (symbol) {
            case INamedTypeSymbol type:
                if (type.BaseType != null) yield return type.BaseType;
                foreach (var contract in type.Interfaces) yield return contract;
                foreach (var parameter in type.TypeParameters) yield return parameter;
                break;
            case IMethodSymbol method:
                yield return method.ReturnType;
                foreach (var parameter in method.Parameters) yield return parameter.Type;
                foreach (var parameter in method.TypeParameters) yield return parameter;
                break;
            case IPropertySymbol property:
                yield return property.Type;
                foreach (var parameter in property.Parameters) yield return parameter.Type;
                break;
            case IFieldSymbol field: yield return field.Type; break;
            case IEventSymbol eventSymbol: yield return eventSymbol.Type; break;
        }
    }
}
