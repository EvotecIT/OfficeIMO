using System.Collections.Immutable;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Xunit;

namespace OfficeIMO.Fluent.Generators.Tests;

public sealed class FluentGeneratorDiagnosticsTests {
    [Fact]
    public void GeneratedSourceCompilesWithValidDocumentationAndNoReflection() {
        var (compilation, diagnostics, generated) = Run("""
            using OfficeIMO.Fluent;
            namespace Sample;
            /// <summary>Run.</summary>
            [Fluent]
            public class Run {
                /// <summary>Bold.</summary>
                public bool? Bold { get; set; }
                /// <summary>Size.</summary>
                public double? Size { get; private set; }
                /// <summary>Options.</summary>
                public Options Options { get; } = new Options();
                /// <summary>A copy of the options.</summary>
                public Options Snapshot => (new Options { Width = Options.Width });
                /// <summary>A copy made by the copy constructor.</summary>
                public Options Copied => new Options(Options);
                /// <summary>A live view over this run.</summary>
                public Margins Margins => new Margins(this);
                private Options? _lazy;
                /// <summary>Options created on first use and kept.</summary>
                public Options Lazy {
                    get {
                        Options Create() { return new Options(); }
                        return _lazy ??= Create();
                    }
                }
                /// <summary>Callback.</summary>
                public System.Func<string, string>? Transform { get; set; }
                [FluentMutator("Size")]
                internal void SetSize(double points, string? unit = null) => Size = points;
            }
            /// <summary>Options.</summary>
            public sealed class Options {
                /// <summary>Creates options.</summary>
                public Options() { }
                /// <summary>Copies options.</summary>
                public Options(Options source) => Width = source.Width;
                /// <summary>Width.</summary>
                public double Width { get; set; }
            }
            /// <summary>Margins of a run.</summary>
            public sealed class Margins {
                private readonly Run _run;
                /// <summary>Wraps a run.</summary>
                public Margins(Run run) => _run = run;
                /// <summary>Left.</summary>
                public double Left { get => _run.Options.Width; set => _run.Options.Width = value; }
            }
            /// <summary>Cell.</summary>
            [Fluent]
            public readonly struct Cell {
                [FluentMutator("Value")]
                internal void SetValue(object? value) { }
            }
            """);

        Assert.Empty(diagnostics);
        Assert.Empty(compilation.GetDiagnostics().Where(d => d.Severity >= DiagnosticSeverity.Warning));
        string source = string.Join("\n", generated);
        Assert.DoesNotContain("System.Reflection", source);
        Assert.DoesNotContain("typeof(", source);
        Assert.DoesNotContain(" Transform(", source);
        Assert.Contains(" Options<TSelf>(", source);
        Assert.DoesNotContain(" Snapshot<TSelf>(", source);
        Assert.DoesNotContain(" Copied<TSelf>(", source);
        Assert.Contains(" Margins<TSelf>(", source);
        Assert.Contains(" Lazy<TSelf>(", source);
    }

    [Theory]
    [InlineData("OIFLUENT001", "[Fluent] public class Run { public bool Bold { get; set; } public Run Bold(bool value) => this; }")]
    [InlineData("OIFLUENT002", "[Fluent] public struct Cell { public int Value { get; set; } }")]
    [InlineData("OIFLUENT003", "[Fluent] public class Run { [FluentMutator(\"Size\")] private void SetSize(double value) { } }")]
    [InlineData("OIFLUENT004", "[Fluent] public class Run<T> { public int Size { get; set; } }")]
    [InlineData("OIFLUENT005", "[Fluent] public class Run { [Fluent(\"Size\")] public int Width { get; set; } public int Size { get; set; } }")]
    public void ReportsShapesItCannotGenerateSafely(string id, string declaration) {
        var (_, diagnostics, _) = Run("using OfficeIMO.Fluent;\nnamespace Sample;\n" + declaration);

        Assert.Contains(diagnostics, diagnostic => diagnostic.Id == id && diagnostic.Severity == DiagnosticSeverity.Error);
    }

    private static (Compilation Compilation, ImmutableArray<Diagnostic> Diagnostics, IReadOnlyList<string> Generated) Run(string source) {
        var references = ((string)AppContext.GetData("TRUSTED_PLATFORM_ASSEMBLIES")!).Split(Path.PathSeparator)
            .Select(path => MetadataReference.CreateFromFile(path));
        var parse = new CSharpParseOptions(LanguageVersion.Latest, DocumentationMode.Diagnose);
        var compilation = CSharpCompilation.Create("FluentSample", [CSharpSyntaxTree.ParseText(source, parse)], references,
            new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary, nullableContextOptions: NullableContextOptions.Enable));
        GeneratorDriver driver = CSharpGeneratorDriver.Create([new FluentGenerator().AsSourceGenerator()], parseOptions: parse);
        driver = driver.RunGeneratorsAndUpdateCompilation(compilation, out var output, out var diagnostics);
        var generated = driver.GetRunResult().GeneratedTrees.Select(tree => tree.ToString()).ToArray();
        return (output, diagnostics, generated);
    }
}
