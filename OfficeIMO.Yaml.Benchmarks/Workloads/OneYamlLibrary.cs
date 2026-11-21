using System.Collections;
using System.Text;
using OfficeIMO.Yaml.Benchmarks.Adapters;
using OfficeIMO.Yaml.Tests.Conformance;
using One = OneYaml;

namespace OfficeIMO.Yaml.Benchmarks.Workloads;

/// <summary>
/// OneYaml in the load, load-and-save and edit-and-save lanes.
/// <list type="bullet">
/// <item>Load: OneYaml's representation is its lossless document, so the tree lane is its untyped deserialization into
/// dictionaries, lists and resolved scalars for every document, as for VYaml. It reads UTF-8 bytes, its native input,
/// encoded once outside timing.</item>
/// <item>Load and save, edit and save: <see cref="One.YamlDocument"/> parsed from the input string and written back as
/// a string, the form the other lanes return.</item>
/// </list>
/// OneYaml has no public event parser and no event emitter, so it has no parse or emit lane.
/// </summary>
internal sealed class OneYamlLibrary : IYamlReferenceLibrary {
    private readonly Dictionary<string, byte[]> _utf8 = new(ReferenceEqualityComparer.Instance);

    public static OneYamlLibrary Instance { get; } = new();

    public string Name => "OneYaml";

    public IYamlConformanceEngine Conformance { get; } = new OneYamlConformanceEngine();

    public int ParseEvents(string yaml) => throw new NotSupportedException("OneYaml has no public event parser.");

    public int LoadTree(string yaml) {
        int count = 0;
        foreach (object? document in One.YamlSerializer.DeserializeAll<object>(Utf8(yaml))) count += Count(document);
        return count;
    }

    public string LoadAndSave(string yaml) {
        using One.YamlDocument document = One.YamlDocument.Parse(yaml);
        return document.ToString();
    }

    // The same value as the other lanes: the first mapping in document order whose first value is a scalar.
    public string EditAndSave(string yaml) {
        using One.YamlDocument document = One.YamlDocument.Parse(yaml);
        if (document.Documents.Count > 0 && FindFirst(document.Documents[0]) is { } scalar) scalar.Set("edited");
        return document.ToString();
    }

    public object CaptureEvents(string yaml) => throw new NotSupportedException("OneYaml has no event emitter.");

    public string Emit(object events) => throw new NotSupportedException("OneYaml has no event emitter.");

    private static One.YamlNode? FindFirst(One.YamlNode node) {
        if (node.Kind == One.YamlNodeKind.Mapping) {
            foreach (KeyValuePair<One.YamlNode, One.YamlNode> entry in node.Entries) {
                if (entry.Value.Kind == One.YamlNodeKind.Scalar) return entry.Value;
                break;
            }
            foreach (KeyValuePair<One.YamlNode, One.YamlNode> entry in node.Entries) {
                if (FindFirst(entry.Key) is { } inKey) return inKey;
                if (FindFirst(entry.Value) is { } inValue) return inValue;
            }
        } else if (node.Kind == One.YamlNodeKind.Sequence) {
            foreach (One.YamlNode item in node.Items) {
                if (FindFirst(item) is { } found) return found;
            }
        }

        return null;
    }

    // Inputs are encoded on first use, which validation does during setup, so timed calls reuse the bytes.
    private byte[] Utf8(string yaml) {
        lock (_utf8) {
            if (!_utf8.TryGetValue(yaml, out byte[]? bytes)) {
                bytes = Encoding.UTF8.GetBytes(yaml);
                _utf8.Add(yaml, bytes);
            }

            return bytes;
        }
    }

    private static int Count(object? value) {
        switch (value) {
            case IDictionary mapping: {
                int count = 1;
                foreach (DictionaryEntry entry in mapping) count += 1 + Count(entry.Value);
                return count;
            }
            case IList sequence: {
                int count = 1;
                foreach (object? item in sequence) count += Count(item);
                return count;
            }
            default:
                return 1;
        }
    }
}
