using System.Text;
using OfficeIMO.Yaml.Tests.Conformance;

namespace OfficeIMO.Yaml.Benchmarks.Workloads;

/// <summary>
/// Setup-time checks that each measured operation does the intended work. They run outside timing and throw, failing
/// the benchmark, when output is wrong. The conformance adapters provide the independent view used for comparison.
/// </summary>
internal static class YamlWorkloadValidation {
    private static readonly UTF8Encoding Utf8 = new(false);

    /// <summary>Event counts match the adapter, and both libraries report the same content.</summary>
    public static void ValidateParse(IYamlReferenceLibrary library, YamlWorkload workload, string[] inputs) {
        IYamlReferenceLibrary other = Other(library);
        foreach (string input in Distinct(inputs)) {
            byte[] bytes = Utf8.GetBytes(input);
            IReadOnlyList<YamlTestEvent> events = library.Conformance.Parse(bytes);
            int count = library.ParseEvents(input);
            // The adapter drops comment events, which the parsers here do not report by default.
            Require(count == events.Count, library, workload, "parsed " + count + " events but the adapter saw " + events.Count);
            int difference = YamlTestEventText.FindContentDifference(other.Conformance.Parse(bytes), events);
            Require(difference < 0, library, workload, "events differ from " + other.Name + " at event " + difference);
        }
    }

    /// <summary>The tree has one node per scalar, collection and alias event, and loads the same values as the other library.</summary>
    public static void ValidateLoad(IYamlReferenceLibrary library, YamlWorkload workload, string[] inputs) {
        IYamlReferenceLibrary other = Other(library);
        foreach (string input in Distinct(inputs)) {
            byte[] bytes = Utf8.GetBytes(input);
            int nodes = library.LoadTree(input);
            int expected = library.Conformance.Parse(bytes).Count(e => e.Kind is YamlTestEventKind.Scalar or YamlTestEventKind.Alias or YamlTestEventKind.MappingStart or YamlTestEventKind.SequenceStart);
            Require(nodes >= expected, library, workload, "loaded " + nodes + " nodes for " + expected + " node events");
            string? difference = YamlJsonEquivalence.FindDifference(other.Conformance.Load(bytes), library.Conformance.Load(bytes));
            Require(difference == null, library, workload, "loaded values differ from " + other.Name + ": " + difference);
        }
    }

    /// <summary>Saving keeps every loaded value.</summary>
    public static void ValidateLoadAndSave(IYamlReferenceLibrary library, YamlWorkload workload, string[] inputs) {
        foreach (string input in Distinct(inputs)) {
            string saved = library.LoadAndSave(input);
            string? difference = YamlJsonEquivalence.FindDifference(library.Conformance.Load(Utf8.GetBytes(input)), library.Conformance.Load(Utf8.GetBytes(saved)));
            Require(difference == null, library, workload, "saved output loads differently: " + difference);
        }
    }

    /// <summary>Emitted output parses back to the same content, anchors and aliases included.</summary>
    public static void ValidateEmit(IYamlReferenceLibrary library, YamlWorkload workload, string[] inputs, object[] captured) {
        string[] distinct = Distinct(inputs).ToArray();
        for (int i = 0; i < distinct.Length; i++) {
            string emitted = library.Emit(captured[Array.IndexOf(inputs, distinct[i])]);
            int difference = YamlTestEventText.FindContentDifference(library.Conformance.Parse(Utf8.GetBytes(distinct[i])), library.Conformance.Parse(Utf8.GetBytes(emitted)));
            Require(difference < 0, library, workload, "emitted output differs at event " + difference);
        }
    }

    private static IYamlReferenceLibrary Other(IYamlReferenceLibrary library) =>
        ReferenceEquals(library, YamlReferenceLibraries.YamlDotNet) ? YamlReferenceLibraries.SharpYaml : YamlReferenceLibraries.YamlDotNet;

    private static IEnumerable<string> Distinct(string[] inputs) => inputs.Distinct(StringComparer.Ordinal);

    private static void Require(bool condition, IYamlReferenceLibrary library, YamlWorkload workload, string message) {
        if (!condition) throw new InvalidOperationException(library.Name + " " + workload + ": " + message + ".");
    }
}
