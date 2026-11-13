using System.Collections;
using System.Reflection;
using System.Text;
using System.Text.Json;
using BenchmarkDotNet.Attributes;
using TomlynReference;

namespace OfficeIMO.Toml.Benchmarks;

/// <summary>The serializer workloads: one corpus file each, mapped to its model.</summary>
public enum TomlSerializerWorkload {
    /// <summary>PyPA's sample <c>pyproject.toml</c> as <see cref="PyProject"/>.</summary>
    SmallConfig,
    /// <summary>uv's workspace <c>Cargo.toml</c> as <see cref="CargoWorkspaceManifest"/>.</summary>
    CargoWorkspace,
    /// <summary>uv's <c>uv.lock</c> as <see cref="UvLock"/>.</summary>
    LockFile
}

/// <summary>The same model mapping in each library, through reflection and through generated metadata.</summary>
public static class TomlSerializerLibraries {
    private static readonly TomlSerializerOptions OfficeIMOReflection = new TomlSerializerOptions { PropertyNamingPolicy = OfficeNamingPolicy.KebabCaseLower };
    private static readonly OfficeIMOBenchmarkContext OfficeIMOGenerated = new OfficeIMOBenchmarkContext(new TomlSerializerOptions { PropertyNamingPolicy = OfficeNamingPolicy.KebabCaseLower });
    private static readonly Tomlyn.TomlSerializerOptions TomlynReflection = new Tomlyn.TomlSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.KebabCaseLower };

    public static Type ModelType(TomlSerializerWorkload workload) => workload switch {
        TomlSerializerWorkload.SmallConfig => typeof(PyProject),
        TomlSerializerWorkload.CargoWorkspace => typeof(CargoWorkspaceManifest),
        _ => typeof(UvLock)
    };

    public static string Input(TomlSerializerWorkload workload) => TomlWorkloads.Get(workload switch {
        TomlSerializerWorkload.SmallConfig => TomlWorkload.SmallConfig,
        TomlSerializerWorkload.CargoWorkspace => TomlWorkload.CargoWorkspace,
        _ => TomlWorkload.LockFile
    })[0];

    public static object ReadOfficeIMO(string toml, Type type) => TomlSerializer.Deserialize(toml, type, OfficeIMOReflection);

    public static object ReadOfficeIMOGenerated(string toml, Type type) => TomlSerializer.Deserialize(toml, type, OfficeIMOGenerated);

    public static object ReadTomlyn(string toml, Type type) => Tomlyn.TomlSerializer.Deserialize(toml, type, TomlynReflection)!;

    public static object ReadTomlynGenerated(string toml, Type type) => Tomlyn.TomlSerializer.Deserialize(toml, type, TomlynBenchmarkContext.Default)!;

    public static string WriteOfficeIMO(object value, Type type) => TomlSerializer.Serialize(value, type, OfficeIMOReflection);

    public static string WriteOfficeIMOGenerated(object value, Type type) => TomlSerializer.Serialize(value, type, OfficeIMOGenerated);

    public static string WriteTomlyn(object value, Type type) => Tomlyn.TomlSerializer.Serialize(value, type, TomlynReflection);

    public static string WriteTomlynGenerated(object value, Type type) => Tomlyn.TomlSerializer.Serialize(value, type, TomlynBenchmarkContext.Default);

    /// <summary>
    /// Checks that every lane reads the same members and values, and that every lane's output reads back to them, so no
    /// lane can win by reading or writing less.
    /// </summary>
    public static void Validate(TomlSerializerWorkload workload) {
        Type type = ModelType(workload);
        string input = Input(workload);
        string expected = Describe(ReadOfficeIMOGenerated(input, type));
        Expect(expected, Describe(ReadOfficeIMO(input, type)), workload, "OfficeIMO reflection read");
        Expect(expected, Describe(ReadTomlyn(input, type)), workload, "Tomlyn read");
        Expect(expected, Describe(ReadTomlynGenerated(input, type)), workload, "Tomlyn generated read");
        object model = ReadOfficeIMOGenerated(input, type);
        Expect(expected, Describe(ReadTomlyn(WriteOfficeIMOGenerated(model, type), type)), workload, "OfficeIMO output read by Tomlyn");
        Expect(expected, Describe(ReadOfficeIMO(WriteOfficeIMO(model, type), type)), workload, "OfficeIMO reflection output");
        // Each library writes the model it read: untyped members hold its own model types.
        object tomlynModel = ReadTomlynGenerated(input, type);
        Expect(expected, Describe(ReadOfficeIMO(WriteTomlyn(tomlynModel, type), type)), workload, "Tomlyn output read by OfficeIMO");
        Expect(expected, Describe(ReadOfficeIMO(WriteTomlynGenerated(tomlynModel, type), type)), workload, "Tomlyn generated output read by OfficeIMO");
    }

    private static void Expect(string expected, string actual, TomlSerializerWorkload workload, string what) {
        if (expected != actual) throw new InvalidOperationException(workload + ": " + what + " differs from OfficeIMO's generated read.");
    }

    /// <summary>
    /// A canonical text of a model: every typed member with its value, dictionaries in key order. Untyped members hold
    /// each library's own model types, so only their presence counts.
    /// </summary>
    public static string Describe(object? value) {
        var text = new StringBuilder();
        if (value is not null) Describe(value, value.GetType(), text);
        return text.ToString();
    }

    private static void Describe(object? value, Type declared, StringBuilder text) {
        if (value is null) {
            text.Append("null;");
            return;
        }
        if (declared == typeof(object)) {
            text.Append("present;");
            return;
        }
        switch (value) {
            case string or bool or long or int or double:
                text.Append(Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture)).Append(';');
                return;
            case IDictionary dictionary: {
                Type valueType = declared.GetGenericArguments()[1];
                text.Append('{');
                foreach (string key in dictionary.Keys.Cast<string>().OrderBy(k => k, StringComparer.Ordinal)) {
                    text.Append(key).Append('=');
                    Describe(dictionary[key], valueType, text);
                }
                text.Append('}');
                return;
            }
            case IEnumerable items: {
                Type itemType = declared.GetGenericArguments()[0];
                text.Append('[');
                foreach (object? item in items) Describe(item, itemType, text);
                text.Append(']');
                return;
            }
        }
        text.Append('(');
        foreach (PropertyInfo property in value.GetType().GetProperties()) {
            text.Append(property.Name).Append(':');
            Describe(property.GetValue(value), Nullable.GetUnderlyingType(property.PropertyType) ?? property.PropertyType, text);
        }
        text.Append(')');
    }
}

/// <summary>Read a corpus file into its model.</summary>
[MemoryDiagnoser]
public class TomlDeserializeBenchmarks {
    private string _input = string.Empty;
    private Type _type = typeof(object);

    [ParamsAllValues]
    public TomlSerializerWorkload Workload { get; set; }

    [GlobalSetup]
    public void Setup() {
        _input = TomlSerializerLibraries.Input(Workload);
        _type = TomlSerializerLibraries.ModelType(Workload);
        TomlSerializerLibraries.Validate(Workload);
    }

    [Benchmark(Baseline = true)]
    public object Tomlyn() => TomlSerializerLibraries.ReadTomlyn(_input, _type);

    [Benchmark]
    public object TomlynGenerated() => TomlSerializerLibraries.ReadTomlynGenerated(_input, _type);

    [Benchmark]
    public object OfficeIMOToml() => TomlSerializerLibraries.ReadOfficeIMO(_input, _type);

    [Benchmark]
    public object OfficeIMOTomlGenerated() => TomlSerializerLibraries.ReadOfficeIMOGenerated(_input, _type);
}

/// <summary>Write the model of a corpus file as TOML. Each library writes the model it read, whose untyped members hold its own model types.</summary>
[MemoryDiagnoser]
public class TomlSerializeBenchmarks {
    private object _model = new object();
    private object _tomlynModel = new object();
    private Type _type = typeof(object);

    [ParamsAllValues]
    public TomlSerializerWorkload Workload { get; set; }

    [GlobalSetup]
    public void Setup() {
        _type = TomlSerializerLibraries.ModelType(Workload);
        _model = TomlSerializerLibraries.ReadOfficeIMOGenerated(TomlSerializerLibraries.Input(Workload), _type);
        _tomlynModel = TomlSerializerLibraries.ReadTomlynGenerated(TomlSerializerLibraries.Input(Workload), _type);
        TomlSerializerLibraries.Validate(Workload);
    }

    [Benchmark(Baseline = true)]
    public int Tomlyn() => TomlSerializerLibraries.WriteTomlyn(_tomlynModel, _type).Length;

    [Benchmark]
    public int TomlynGenerated() => TomlSerializerLibraries.WriteTomlynGenerated(_tomlynModel, _type).Length;

    [Benchmark]
    public int OfficeIMOToml() => TomlSerializerLibraries.WriteOfficeIMO(_model, _type).Length;

    [Benchmark]
    public int OfficeIMOTomlGenerated() => TomlSerializerLibraries.WriteOfficeIMOGenerated(_model, _type).Length;
}
