// Compiles the C# code blocks in the website documentation against the current source.
// Build/Test-DocsCodeExamples.ps1 generates the project that references the engines and runs this.
using System.Globalization;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.Diagnostics;

if (args.Length < 1) {
    Console.Error.WriteLine("Usage: DocsExamples <docs-root> [<docs-root> ...]");
    return 2;
}

string[] implicitUsings = [
    "System", "System.Collections.Generic", "System.IO", "System.Linq",
    "System.Net.Http", "System.Threading", "System.Threading.Tasks"
];
SyntaxTree globalUsings = CSharpSyntaxTree.ParseText(
    string.Concat(implicitUsings.Select(ns => $"global using global::{ns};\n")),
    new CSharpParseOptions(LanguageVersion.Preview), path: "GlobalUsings.g.cs");

List<MetadataReference> references = LoadReferences();
// Source generators (OfficeIMO.Yaml/Toml serializer contexts) run the way they would in a consuming project.
ISourceGenerator[] generators = Directory.EnumerateFiles(AppContext.BaseDirectory, "*.SourceGeneration.dll")
    .SelectMany(path => new AnalyzerFileReference(path, new GeneratorLoader()).GetGenerators(LanguageNames.CSharp))
    .ToArray();
var parseOptions = new CSharpParseOptions(LanguageVersion.Preview);
int blocks = 0, skipped = 0, failed = 0;

foreach (string root in args) {
    foreach (string file in Directory.EnumerateFiles(root, "*.md", SearchOption.AllDirectories).Order(StringComparer.Ordinal)) {
        foreach (CodeBlock block in ExtractBlocks(file)) {
            if (block.Skip) {
                skipped++;
                continue;
            }

            blocks++;
            IReadOnlyList<Diagnostic> errors = Compile(block, references, generators, globalUsings, parseOptions);
            if (errors.Count == 0) continue;

            failed++;
            foreach (Diagnostic error in errors.Take(5)) {
                int line = block.FirstLine + error.Location.GetLineSpan().StartLinePosition.Line;
                Console.WriteLine($"{Path.GetRelativePath(Directory.GetCurrentDirectory(), file)}:{line}: {error.Id} {error.GetMessage(CultureInfo.InvariantCulture)}");
            }
        }
    }
}

Console.WriteLine($"Compiled {blocks} C# blocks: {blocks - failed} passed, {failed} failed, {skipped} skipped by marker.");
return failed == 0 ? 0 : 1;

static IReadOnlyList<Diagnostic> Compile(CodeBlock block, List<MetadataReference> references, ISourceGenerator[] generators, SyntaxTree globalUsings, CSharpParseOptions parseOptions) {
    SyntaxTree tree = CSharpSyntaxTree.ParseText(block.Code, parseOptions, path: block.File);
    bool hasTopLevelStatements = tree.GetCompilationUnitRoot().Members.OfType<Microsoft.CodeAnalysis.CSharp.Syntax.GlobalStatementSyntax>().Any();
    var options = new CSharpCompilationOptions(
        hasTopLevelStatements ? OutputKind.ConsoleApplication : OutputKind.DynamicallyLinkedLibrary,
        nullableContextOptions: NullableContextOptions.Enable,
        allowUnsafe: false);
    Compilation compilation = CSharpCompilation.Create("DocsExample", [globalUsings, tree], references, options);
    if (generators.Length > 0) {
        CSharpGeneratorDriver.Create(generators, parseOptions: parseOptions)
            .RunGeneratorsAndUpdateCompilation(compilation, out compilation, out _);
    }
    using var output = new MemoryStream();
    var result = compilation.Emit(output);
    return result.Diagnostics.Where(d => d.Severity == DiagnosticSeverity.Error).ToList();
}

static List<MetadataReference> LoadReferences() {
    var paths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    if (AppContext.GetData("TRUSTED_PLATFORM_ASSEMBLIES") is string tpa) {
        foreach (string path in tpa.Split(Path.PathSeparator)) paths.Add(path);
    }
    string baseDirectory = AppContext.BaseDirectory;
    foreach (string path in Directory.EnumerateFiles(baseDirectory, "*.dll")) {
        if (IsManagedAssembly(path) && !path.EndsWith(".SourceGeneration.dll", StringComparison.OrdinalIgnoreCase)) paths.Add(path);
    }
    // One reference per simple name keeps duplicate framework copies from causing ambiguities.
    return paths
        .GroupBy(p => Path.GetFileNameWithoutExtension(p), StringComparer.OrdinalIgnoreCase)
        .Select(g => (MetadataReference)MetadataReference.CreateFromFile(g.First()))
        .ToList();
}

static bool IsManagedAssembly(string path) {
    try {
        AssemblyName.GetAssemblyName(path);
        return true;
    } catch (BadImageFormatException) {
        return false;
    }
}

static IEnumerable<CodeBlock> ExtractBlocks(string file) {
    string[] lines = File.ReadAllLines(file);
    bool skipNext = false;
    for (int i = 0; i < lines.Length; i++) {
        string trimmed = lines[i].Trim();
        if (trimmed.StartsWith("<!--", StringComparison.Ordinal) && trimmed.Contains("docs-check: skip", StringComparison.Ordinal)) {
            skipNext = true;
            continue;
        }

        Match fence = Regex.Match(lines[i], @"^(\s*)(`{3,}|~{3,})\s*(csharp|cs|c#)\s*$", RegexOptions.IgnoreCase);
        if (!fence.Success) {
            if (trimmed.Length > 0 && !trimmed.StartsWith("```", StringComparison.Ordinal)) skipNext = skipNext && trimmed.StartsWith("<!--", StringComparison.Ordinal);
            continue;
        }

        string marker = fence.Groups[2].Value;
        var code = new StringBuilder();
        int start = i + 2;
        for (i++; i < lines.Length && !lines[i].TrimStart().StartsWith(marker, StringComparison.Ordinal); i++) {
            code.AppendLine(lines[i]);
        }
        yield return new CodeBlock(file, start, code.ToString(), skipNext);
        skipNext = false;
    }
}

internal sealed class GeneratorLoader : IAnalyzerAssemblyLoader {
    public void AddDependencyLocation(string fullPath) { }
    public Assembly LoadFromPath(string fullPath) => Assembly.LoadFrom(fullPath);
}

internal sealed record CodeBlock(string File, int FirstLine, string Code, bool Skip);
