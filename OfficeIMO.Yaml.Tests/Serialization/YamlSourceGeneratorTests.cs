using System.Collections.Immutable;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using OfficeIMO.Yaml.SourceGeneration;
using Xunit;

namespace OfficeIMO.Yaml.Tests.Serialization;

// Contract: a shape the serializer cannot handle is a build error at the context, never a runtime fallback. Each case
// must produce exactly its diagnostic, so a harness or reference problem cannot pass for the expected error.
public class YamlSourceGeneratorTests {
    private const string Usings = "using System; using System.Collections; using System.Collections.Generic; using OfficeIMO.Yaml;\n";

    [Theory]
    [InlineData("YAML001", "[YamlSerializable(typeof(int))] public class Context : YamlSerializerContext { }")]
    [InlineData("YAML002", "public class Model { public Action? Callback { get; set; } } [YamlSerializable(typeof(Model))] public partial class Context : YamlSerializerContext { }")]
    [InlineData("YAML002", "public class Model { public int[,] Grid { get; set; } = new int[0, 0]; } [YamlSerializable(typeof(Model))] public partial class Context : YamlSerializerContext { }")]
    [InlineData("YAML002", "public class Model { public ArrayList Items { get; set; } = new(); } [YamlSerializable(typeof(Model))] public partial class Context : YamlSerializerContext { }")]
    [InlineData("YAML002", "public abstract class Shape { } public class Model { public Shape? Shape { get; set; } } [YamlSerializable(typeof(Model))] public partial class Context : YamlSerializerContext { }")]
    [InlineData("YAML002", "public class Key { } public class Model { public Dictionary<Key, int> Map { get; set; } = new(); } [YamlSerializable(typeof(Model))] public partial class Context : YamlSerializerContext { }")]
    [InlineData("YAML002", "public class Model { public IntPtr Handle { get; set; } } [YamlSerializable(typeof(Model))] public partial class Context : YamlSerializerContext { }")]
    [InlineData("YAML003", "public class Model { public Model(int a) { } public Model(string b) { } public int A { get; set; } } [YamlSerializable(typeof(Model))] public partial class Context : YamlSerializerContext { }")]
    [InlineData("YAML003", "public class Model { public Model(int count) { } public int Total { get; set; } } [YamlSerializable(typeof(Model))] public partial class Context : YamlSerializerContext { }")]
    [InlineData("YAML003", "public class Model { [YamlConstructor] public Model() { } [YamlConstructor] public Model(int a) { A = a; } public int A { get; set; } } [YamlSerializable(typeof(Model))] public partial class Context : YamlSerializerContext { }")]
    [InlineData("YAML004", "public class Model { [YamlPropertyName(\"x\")] public int A { get; set; } [YamlPropertyName(\"x\")] public int B { get; set; } } [YamlSerializable(typeof(Model))] public partial class Context : YamlSerializerContext { }")]
    [InlineData("YAML005", "public class Other { } [YamlDerivedType(typeof(Other), \"o\")] public class Model { } [YamlSerializable(typeof(Model))] public partial class Context : YamlSerializerContext { }")]
    [InlineData("YAML005", "public class A : Model { } public class B : Model { } [YamlDerivedType(typeof(A), \"same\")] [YamlDerivedType(typeof(B), \"same\")] public class Model { } [YamlSerializable(typeof(Model))] public partial class Context : YamlSerializerContext { }")]
    public void UnsupportedShapesAreBuildErrors(string expectedId, string source) {
        ImmutableArray<Diagnostic> diagnostics = Generate(Usings + source);
        Assert.True(diagnostics.Length == 1 && diagnostics[0].Id == expectedId,
            "Expected only " + expectedId + ", got: " + string.Join("; ", diagnostics.Select(d => d.Id + " " + d.GetMessage())));
        Assert.Equal(DiagnosticSeverity.Error, diagnostics[0].Severity);
    }

    private static ImmutableArray<Diagnostic> Generate(string source) {
        var references = new List<MetadataReference>();
        foreach (System.Reflection.Assembly assembly in AppDomain.CurrentDomain.GetAssemblies()) {
            if (!assembly.IsDynamic && assembly.Location.Length > 0) references.Add(MetadataReference.CreateFromFile(assembly.Location));
        }
        references.Add(MetadataReference.CreateFromFile(typeof(YamlSerializer).Assembly.Location));
        CSharpCompilation compilation = CSharpCompilation.Create("Generated", new[] { CSharpSyntaxTree.ParseText(source, new CSharpParseOptions(LanguageVersion.Latest)) },
            references, new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary, nullableContextOptions: NullableContextOptions.Enable));
        GeneratorDriver driver = CSharpGeneratorDriver.Create(new YamlSourceGenerator()).RunGeneratorsAndUpdateCompilation(compilation, out _, out ImmutableArray<Diagnostic> diagnostics);
        return diagnostics;
    }
}
