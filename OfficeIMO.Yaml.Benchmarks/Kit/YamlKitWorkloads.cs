using System.Collections;
using System.Globalization;
using System.Numerics;
using System.Text;
using System.Text.Json;
using OfficeIMO.Yaml.Tests.Conformance;

namespace OfficeIMO.Yaml.Benchmarks.Kit;

/// <summary>The tasks of the oneyaml-kit benchmark, defined in its <c>docs/benchmarks.md</c>.</summary>
public enum YamlKitTask {
    /// <summary>Input text to the workload's typed model.</summary>
    TypedLoad,
    /// <summary>Input text to the library's untyped values, every document of a stream.</summary>
    GenericLoad,
    /// <summary>The library's own untyped values, read before timing, to YAML text.</summary>
    Dump,
    /// <summary>Input text to a lossless document.</summary>
    DocumentParse,
    /// <summary>A lossless document parsed before timing, written back unedited.</summary>
    UnchangedEmit,
    /// <summary>A lossless document parsed before timing, edited by a fixed script of replace and remove steps, written back.</summary>
    EditScript
}

/// <summary>One workload prepared by the kit's <c>bin/bench workloads</c>: input, expected values and edit script.</summary>
internal sealed class YamlKitWorkload {
    private static readonly UTF8Encoding Utf8NoBom = new(false);

    private YamlKitWorkload(string directory, JsonElement meta) {
        Id = meta.GetProperty("id").GetString()!;
        Documents = meta.GetProperty("documents").GetInt32();
        TypedModel = meta.TryGetProperty("typed_model", out JsonElement typed) && typed.ValueKind == JsonValueKind.String ? typed.GetString() : null;
        Tasks = meta.GetProperty("tasks").EnumerateArray().Select(t => t.GetString()!).Select(ParseTask).Where(t => t != null).Select(t => t!.Value).ToHashSet();
        Bytes = File.ReadAllBytes(Path.Combine(directory, "input.yaml"));
        Text = Utf8NoBom.GetString(Bytes);
        string expected = File.ReadAllText(Path.Combine(directory, "expected.json"));
        Expected = Documents == 1 ? new[] { YamlJsonReader.Read(expected) } : YamlJsonReader.Read(expected).Items;
        string merged = Path.Combine(directory, "expected.merged.json");
        if (File.Exists(merged)) {
            string json = File.ReadAllText(merged);
            ExpectedMerged = Documents == 1 ? new[] { YamlJsonReader.Read(json) } : YamlJsonReader.Read(json).Items;
        }

        string edit = Path.Combine(directory, "edit.json");
        if (File.Exists(edit)) {
            using JsonDocument script = JsonDocument.Parse(File.ReadAllBytes(edit));
            EditOps = script.RootElement.GetProperty("ops").EnumerateArray().Select(op => new YamlKitEdit(
                op.GetProperty("op").GetString()!,
                op.GetProperty("path").GetString()!,
                op.TryGetProperty("value", out JsonElement value) ? Value(value) : null)).ToArray();
            EditExpected = File.ReadAllBytes(Path.Combine(directory, "edit.expected.yaml"));
        }
    }

    public string Id { get; }

    public int Documents { get; }

    public string? TypedModel { get; }

    public IReadOnlyCollection<YamlKitTask> Tasks { get; }

    public byte[] Bytes { get; }

    public string Text { get; }

    public IReadOnlyList<YamlConformanceNode> Expected { get; }

    /// <summary>The value when <c>&lt;&lt;</c> merge keys are applied, a YAML 1.1 extension the kit also accepts.</summary>
    public IReadOnlyList<YamlConformanceNode>? ExpectedMerged { get; }

    public IReadOnlyList<YamlKitEdit> EditOps { get; } = Array.Empty<YamlKitEdit>();

    public byte[]? EditExpected { get; }

    public override string ToString() => Id;

    /// <summary>
    /// The prepared workloads: <c>OFFICEIMO_YAML_KIT_WORKLOADS</c>, or <c>oneyaml-kit-workloads</c> under the temporary
    /// folder, holding the kit's <c>workloads</c> output (an <c>index.json</c> and one folder per workload).
    /// </summary>
    public static string Root => Environment.GetEnvironmentVariable("OFFICEIMO_YAML_KIT_WORKLOADS") is { Length: > 0 } root
        ? root
        : Path.Combine(Path.GetTempPath(), "oneyaml-kit-workloads");

    private static IReadOnlyList<YamlKitWorkload>? _all;

    public static IReadOnlyList<YamlKitWorkload> All {
        get {
            if (_all != null) return _all;
            string index = Path.Combine(Root, "index.json");
            if (!File.Exists(index)) throw new InvalidOperationException("No kit workloads at " + Root + "; run the kit's 'bin/bench workloads --tier full --out <dir>' and set OFFICEIMO_YAML_KIT_WORKLOADS to <dir>/workloads.");
            using JsonDocument document = JsonDocument.Parse(File.ReadAllBytes(index));
            var workloads = new List<YamlKitWorkload>();
            foreach (JsonElement entry in document.RootElement.GetProperty("workloads").EnumerateArray()) {
                string directory = Path.Combine(Root, entry.GetProperty("id").GetString()!);
                using JsonDocument meta = JsonDocument.Parse(File.ReadAllBytes(Path.Combine(directory, "workload.json")));
                workloads.Add(new YamlKitWorkload(directory, meta.RootElement));
            }

            return _all = workloads;
        }
    }

    public static YamlKitWorkload Get(string id) => All.FirstOrDefault(w => w.Id == id) ?? throw new ArgumentException("Unknown kit workload " + id + ".", nameof(id));

    /// <summary>Whether a loaded value is the workload's value, plain or with merge keys applied.</summary>
    public string? Difference(IReadOnlyList<YamlConformanceNode> actual) {
        string? plain = YamlJsonEquivalence.FindDifference(Expected, actual);
        if (plain == null || ExpectedMerged == null) return plain;
        return YamlJsonEquivalence.FindDifference(ExpectedMerged, actual) == null ? null : plain;
    }

    private static YamlKitTask? ParseTask(string name) => name switch {
        "typed-load" => YamlKitTask.TypedLoad,
        "generic-load" => YamlKitTask.GenericLoad,
        "dump" => YamlKitTask.Dump,
        "document-parse" => YamlKitTask.DocumentParse,
        "unchanged-emit" => YamlKitTask.UnchangedEmit,
        "edit-script" => YamlKitTask.EditScript,
        _ => null
    };

    private static object? Value(JsonElement value) => value.ValueKind switch {
        JsonValueKind.Null => null,
        JsonValueKind.True => true,
        JsonValueKind.False => false,
        JsonValueKind.String => value.GetString(),
        JsonValueKind.Number => value.TryGetInt64(out long integer) ? (object)integer : value.GetDouble(),
        _ => throw new NotSupportedException("Edit values are scalars in the kit's scripts.")
    };

    /// <summary>Plain .NET values (dictionaries, lists, scalars) as the harness's comparable form.</summary>
    public static YamlConformanceNode ToNode(object? value, int depth = 0) {
        if (depth > 1000) throw new InvalidOperationException("Nesting is too deep or recursive.");
        switch (value) {
            case null: return YamlConformanceNode.Null;
            case bool flag: return YamlConformanceNode.Boolean(flag);
            case string text: return YamlConformanceNode.String(text);
            case int number: return YamlConformanceNode.Integer(number.ToString(CultureInfo.InvariantCulture));
            case long number: return YamlConformanceNode.Integer(number.ToString(CultureInfo.InvariantCulture));
            case ulong number: return YamlConformanceNode.Integer(number.ToString(CultureInfo.InvariantCulture));
            case BigInteger number: return YamlConformanceNode.Integer(number.ToString(CultureInfo.InvariantCulture));
            case decimal number: return YamlConformanceNode.Float(number.ToString(CultureInfo.InvariantCulture));
            case double number: return YamlConformanceNode.Float(Float(number));
            case float number: return YamlConformanceNode.Float(Float(number));
            case IDictionary mapping: {
                var entries = new List<KeyValuePair<YamlConformanceNode, YamlConformanceNode>>(mapping.Count);
                foreach (DictionaryEntry entry in mapping) entries.Add(new(ToNode(entry.Key, depth + 1), ToNode(entry.Value, depth + 1)));
                return YamlConformanceNode.Mapping(entries);
            }
            case IList sequence: {
                var items = new List<YamlConformanceNode>(sequence.Count);
                foreach (object? item in sequence) items.Add(ToNode(item, depth + 1));
                return YamlConformanceNode.Sequence(items);
            }
            default:
                throw new NotSupportedException("Unexpected value " + value.GetType().FullName + ".");
        }
    }

    private static string Float(double value) =>
        double.IsPositiveInfinity(value) ? "Infinity" : double.IsNegativeInfinity(value) ? "-Infinity" : double.IsNaN(value) ? "NaN" : value.ToString("R", CultureInfo.InvariantCulture);

    /// <summary>The typed models as the comparable form, with the members the expected value holds.</summary>
    public static YamlConformanceNode ToNode(List<KitRecord> records) => YamlConformanceNode.Sequence(records.Select(r => Map(
        ("id", r.Id), ("name", r.Name), ("email", r.Email), ("active", r.Active), ("score", r.Score),
        ("tags", r.Tags.Cast<object?>().ToList()),
        ("address", Map(("city", r.Address.City), ("zip", r.Address.Zip), ("geo", r.Address.Geo == null ? null : Map(("lat", r.Address.Geo.Lat), ("lng", r.Address.Geo.Lng))))),
        ("note", r.Note))).Select(m => ToNode(m)).ToList());

    public static YamlConformanceNode ToNode(KitWorkflow workflow) {
        var jobs = new Dictionary<string, object?>();
        foreach (KeyValuePair<string, KitJob> entry in workflow.Jobs) {
            KitJob job = entry.Value;
            var map = new Dictionary<string, object?> { ["runs-on"] = job.RunsOn };
            if (job.Strategy != null) map["strategy"] = Map(("matrix", Map(("runs-on", job.Strategy.Matrix.RunsOn.Cast<object?>().ToList()))));
            if (job.Container != null) map["container"] = job.Container;
            if (job.Services != null) map["services"] = job.Services.ToDictionary(s => s.Key, s => (object?)Map(("image", s.Value.Image), ("ports", s.Value.Ports.Cast<object?>().ToList())));
            if (job.Env != null) map["env"] = job.Env.ToDictionary(e => e.Key, e => (object?)e.Value);
            map["steps"] = job.Steps.Select(Step).Cast<object?>().ToList();
            jobs[entry.Key] = map;
        }

        Dictionary<string, object?> push = Map(("branches", workflow.On.Push.Branches.Cast<object?>().ToList()));
        Dictionary<string, object?> on = Map(("pull_request", workflow.On.PullRequest), ("push", push));
        return ToNode(Map(("name", workflow.Name), ("on", on), ("jobs", jobs)));
    }

    private static Dictionary<string, object?> Step(KitStep step) {
        var map = new Dictionary<string, object?>();
        if (step.Id != null) map["id"] = step.Id;
        if (step.Name != null) map["name"] = step.Name;
        if (step.Condition != null) map["if"] = step.Condition;
        if (step.Uses != null) map["uses"] = step.Uses;
        if (step.Run != null) map["run"] = step.Run;
        if (step.Shell != null) map["shell"] = step.Shell;
        if (step.With != null) map["with"] = step.With.ToDictionary(w => w.Key, w => w.Value);
        return map;
    }

    private static Dictionary<string, object?> Map(params (string Key, object? Value)[] entries) {
        var map = new Dictionary<string, object?>(entries.Length);
        foreach ((string key, object? value) in entries) map[key] = value;
        return map;
    }
}

/// <summary>One step of an edit script: <c>replace</c> or <c>remove</c> at an RFC 6901 pointer.</summary>
internal sealed record YamlKitEdit(string Op, string Pointer, object? Value) {
    /// <summary>The pointer's reference tokens, unescaped.</summary>
    public string[] Tokens => Pointer.Length == 0 ? Array.Empty<string>() : Pointer.Substring(1).Split('/').Select(t => t.Replace("~1", "/").Replace("~0", "~")).ToArray();
}
