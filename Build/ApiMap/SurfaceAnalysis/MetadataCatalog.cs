using System.Text.Json;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;

namespace OfficeIMO.Build.ApiMap;

/// <summary>Loads the built public surface without compiling implementation bodies.</summary>
internal sealed class MetadataCatalog {
    internal CSharpCompilation Compilation { get; }
    internal IReadOnlyDictionary<string, IAssemblySymbol> Assemblies { get; }
    internal IReadOnlyList<INamedTypeSymbol> Types { get; }

    internal MetadataCatalog(string root, IEnumerable<string> packages) {
        var paths = ((string)AppContext.GetData("TRUSTED_PLATFORM_ASSEMBLIES")!).Split(Path.PathSeparator)
            .GroupBy(path => Path.GetFileNameWithoutExtension(path), StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group.First(), StringComparer.OrdinalIgnoreCase);
        var targets = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (string package in packages) {
            string path = Path.Combine(root, package, "bin", "Release", "net10.0", package + ".dll");
            if (!File.Exists(path)) throw new FileNotFoundException("Build the net10.0 Release surface before analysis.", path);
            paths[package] = path;
            targets[package] = path;
            AddCompileDependencies(root, package, paths);
        }
        var references = paths.Values.Distinct(StringComparer.Ordinal).Select(path => MetadataReference.CreateFromFile(path)).ToArray();
        Compilation = CSharpCompilation.Create("OfficeIMO.SurfaceAnalysis", references: references,
            options: new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary));
        Assemblies = targets.ToDictionary(item => item.Key, item =>
            (IAssemblySymbol)Compilation.GetAssemblyOrModuleSymbol(references.Single(reference => reference.FilePath == item.Value))!, StringComparer.Ordinal);
        Types = Assemblies.Values.SelectMany(assembly => PublicTypes(assembly.GlobalNamespace)).ToArray();
    }

    private static void AddCompileDependencies(string root, string package, Dictionary<string, string> paths) {
        string path = Path.Combine(root, package, "obj", "project.assets.json");
        if (!File.Exists(path)) return;
        using var assets = JsonDocument.Parse(File.ReadAllText(path));
        if (!assets.RootElement.GetProperty("targets").TryGetProperty("net10.0", out var target)) return;
        string cache = assets.RootElement.GetProperty("packageFolders").EnumerateObject().First().Name;
        foreach (var dependency in target.EnumerateObject()) {
            if (!dependency.Value.TryGetProperty("compile", out var compile)) continue;
            foreach (var item in compile.EnumerateObject()) {
                string assembly = Path.Combine(cache, dependency.Name.ToLowerInvariant(), item.Name);
                string name = Path.GetFileNameWithoutExtension(assembly);
                if (assembly.EndsWith(".dll", StringComparison.Ordinal) && File.Exists(assembly) && !paths.ContainsKey(name)) paths[name] = assembly;
            }
        }
    }

    private static IEnumerable<INamedTypeSymbol> PublicTypes(INamespaceSymbol scope) {
        foreach (var child in scope.GetNamespaceMembers()) foreach (var type in PublicTypes(child)) yield return type;
        foreach (var type in scope.GetTypeMembers()) foreach (var visible in PublicTypes(type)) yield return visible;
    }

    private static IEnumerable<INamedTypeSymbol> PublicTypes(INamedTypeSymbol type) {
        if (type.DeclaredAccessibility != Accessibility.Public) yield break;
        yield return type;
        foreach (var nested in type.GetTypeMembers()) foreach (var visible in PublicTypes(nested)) yield return visible;
    }
}
