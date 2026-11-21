using System.Collections.Immutable;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using OfficeIMO.Toml.SourceGeneration;
using Xunit;

namespace OfficeIMO.Toml.Tests.Serialization;

// Contract: a shape the serializer cannot handle is a build error at the context, never a runtime fallback. Each case
// must produce exactly its diagnostic, so a harness or reference problem cannot pass for the expected error.
public class TomlSourceGeneratorTests {
    private const string Usings = "using System; using System.Collections; using System.Collections.Generic; using OfficeIMO; using OfficeIMO.Toml;\n";

    [Theory]
    [InlineData("TOML001", "[TomlSerializable(typeof(int))] public class Context : TomlSerializerContext { }")]
    [InlineData("TOML002", "public class Model { public Action? Callback { get; set; } } [TomlSerializable(typeof(Model))] public partial class Context : TomlSerializerContext { }")]
    [InlineData("TOML002", "public class Model { public int[,] Grid { get; set; } = new int[0, 0]; } [TomlSerializable(typeof(Model))] public partial class Context : TomlSerializerContext { }")]
    [InlineData("TOML002", "public class Model { public ArrayList Items { get; set; } = new(); } [TomlSerializable(typeof(Model))] public partial class Context : TomlSerializerContext { }")]
    [InlineData("TOML002", "public abstract class Shape { } public class Model { public Shape? Shape { get; set; } } [TomlSerializable(typeof(Model))] public partial class Context : TomlSerializerContext { }")]
    [InlineData("TOML002", "public class Key { } public class Model { public Dictionary<Key, int> Map { get; set; } = new(); } [TomlSerializable(typeof(Model))] public partial class Context : TomlSerializerContext { }")]
    [InlineData("TOML002", "public class Model { public IntPtr Handle { get; set; } } [TomlSerializable(typeof(Model))] public partial class Context : TomlSerializerContext { }")]
    [InlineData("TOML003", "public class Model { public Model(int a) { } public Model(string b) { } public int A { get; set; } } [TomlSerializable(typeof(Model))] public partial class Context : TomlSerializerContext { }")]
    [InlineData("TOML003", "public class Model { public Model(int count) { } public int Total { get; set; } } [TomlSerializable(typeof(Model))] public partial class Context : TomlSerializerContext { }")]
    [InlineData("TOML003", "public class Model { [OfficeConstructor] public Model() { } [OfficeConstructor] public Model(int a) { A = a; } public int A { get; set; } } [TomlSerializable(typeof(Model))] public partial class Context : TomlSerializerContext { }")]
    [InlineData("TOML004", "public class Model { [OfficePropertyName(\"x\")] public int A { get; set; } [OfficePropertyName(\"x\")] public int B { get; set; } } [TomlSerializable(typeof(Model))] public partial class Context : TomlSerializerContext { }")]
    [InlineData("TOML005", "public class Other { } [OfficeDerivedType(typeof(Other), \"o\")] public class Model { } [TomlSerializable(typeof(Model))] public partial class Context : TomlSerializerContext { }")]
    [InlineData("TOML005", "public class A : Model { } public class B : Model { } [OfficeDerivedType(typeof(A), \"same\")] [OfficeDerivedType(typeof(B), \"same\")] public class Model { } [TomlSerializable(typeof(Model))] public partial class Context : TomlSerializerContext { }")]
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
        references.Add(MetadataReference.CreateFromFile(typeof(TomlSerializer).Assembly.Location));
        CSharpCompilation compilation = CSharpCompilation.Create("Generated", new[] { CSharpSyntaxTree.ParseText(source, new CSharpParseOptions(LanguageVersion.Latest)) },
            references, new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary, nullableContextOptions: NullableContextOptions.Enable));
        GeneratorDriver driver = CSharpGeneratorDriver.Create(new TomlSourceGenerator()).RunGeneratorsAndUpdateCompilation(compilation, out _, out ImmutableArray<Diagnostic> diagnostics);
        return diagnostics;
    }
}
