using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;

namespace OfficeIMO.Build.ApiMap;

internal sealed record SurfaceLocation(string File, int Line);

/// <summary>Matches metadata declarations to their source, retaining all partial locations.</summary>
internal sealed class SourceLocationIndex {
    private sealed record Entry(string Type, string Name, int Arity, string[] Parameters, SyntaxNode Syntax);
    private readonly Dictionary<string, List<Entry>> _types = new(StringComparer.Ordinal);

    internal SourceLocationIndex(string root, IEnumerable<string> files, ISet<string> packages) {
        var options = new CSharpParseOptions(LanguageVersion.Preview, preprocessorSymbols:
            ["NET", "NET10_0", "NET10_0_OR_GREATER", "NET8_0_OR_GREATER", "NET6_0_OR_GREATER"]);
        foreach (string file in files.Where(file => file.EndsWith(".cs", StringComparison.Ordinal) && packages.Contains(file.Split('/')[0]))) {
            string path = Path.Combine(root, file);
            if (!File.Exists(path)) continue;
            var syntax = CSharpSyntaxTree.ParseText(File.ReadAllText(path), options, file).GetRoot();
            foreach (var type in syntax.DescendantNodes().OfType<BaseTypeDeclarationSyntax>()) {
                string id = TypeId(type);
                if (!_types.TryGetValue(id, out var entries)) _types[id] = entries = [];
                entries.Add(new Entry(id, "", 0, [], type));
                if (type is not TypeDeclarationSyntax declaration) continue;
                foreach (var member in declaration.Members) {
                    string name = member switch {
                        MethodDeclarationSyntax method => method.Identifier.ValueText,
                        ConstructorDeclarationSyntax => ".ctor",
                        PropertyDeclarationSyntax property => property.Identifier.ValueText,
                        IndexerDeclarationSyntax => "this[]",
                        FieldDeclarationSyntax field => field.Declaration.Variables.First().Identifier.ValueText,
                        EventFieldDeclarationSyntax field => field.Declaration.Variables.First().Identifier.ValueText,
                        _ => ""
                    };
                    if (name.Length == 0) continue;
                    var parameters = member is BaseMethodDeclarationSyntax methodDeclaration ? methodDeclaration.ParameterList.Parameters
                        : member is IndexerDeclarationSyntax indexer ? indexer.ParameterList.Parameters : default;
                    entries.Add(new Entry(id, name, member is MethodDeclarationSyntax generic ? generic.TypeParameterList?.Parameters.Count ?? 0 : 0,
                        parameters.Select(parameter => SimpleType(parameter.Type?.ToString() ?? "")).ToArray(), member));
                }
            }
        }
    }

    internal IReadOnlyList<SurfaceLocation> Find(ISymbol symbol) {
        INamedTypeSymbol type = symbol is INamedTypeSymbol named ? named : symbol.ContainingType;
        if (!_types.TryGetValue(TypeId(type), out var entries)) return [];
        IEnumerable<Entry> matches;
        if (symbol is INamedTypeSymbol) matches = entries.Where(entry => entry.Name.Length == 0);
        else {
            string name = symbol is IMethodSymbol { MethodKind: MethodKind.Constructor } ? ".ctor" : symbol.Name;
            int arity = symbol is IMethodSymbol method ? method.Arity : 0;
            var parameters = symbol is IMethodSymbol methodSymbol ? methodSymbol.Parameters
                : symbol is IPropertySymbol property ? property.Parameters : default;
            string[] types = parameters.IsDefault ? [] : parameters.Select(parameter => SimpleType(parameter.Type.ToDisplayString(SymbolDisplayFormat.MinimallyQualifiedFormat))).ToArray();
            var candidates = entries.Where(entry => entry.Name == name && entry.Arity == arity && entry.Parameters.Length == types.Length).ToArray();
            matches = candidates.Length == 1 ? candidates : candidates.Where(entry => entry.Parameters.SequenceEqual(types, StringComparer.Ordinal));
        }
        return matches.Select(entry => new SurfaceLocation(entry.Syntax.SyntaxTree.FilePath,
            entry.Syntax.GetLocation().GetLineSpan().StartLinePosition.Line + 1)).Distinct().ToArray();
    }

    private static string TypeId(INamedTypeSymbol type) => type.GetDocumentationCommentId() ?? type.ToDisplayString();

    private static string TypeId(BaseTypeDeclarationSyntax type) {
        var parts = type.AncestorsAndSelf().OfType<TypeDeclarationSyntax>().Reverse()
            .Select(parent => parent.Identifier.ValueText + (parent.TypeParameterList is { Parameters.Count: > 0 } parameters ? "`" + parameters.Parameters.Count : "")).ToList();
        if (type is EnumDeclarationSyntax) parts.Add(type.Identifier.ValueText);
        string scope = string.Join(".", type.Ancestors().OfType<BaseNamespaceDeclarationSyntax>().Reverse().Select(parent => parent.Name.ToString()));
        return "T:" + (scope.Length == 0 ? "" : scope + ".") + string.Join(".", parts);
    }

    private static string SimpleType(string name) => System.Text.RegularExpressions.Regex.Replace(name, @"\b(?:[A-Za-z_]\w*\.)+", "")
        .Replace(" ", "").Replace("?", "").Replace("System.String", "string").Replace("System.Int32", "int");
}
