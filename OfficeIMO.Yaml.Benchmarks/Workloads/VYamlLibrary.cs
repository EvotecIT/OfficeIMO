using System.Text;
using OfficeIMO.Yaml.Benchmarks.Adapters;
using OfficeIMO.Yaml.Tests.Conformance;
using VYaml.Parser;
using VYaml.Serialization;

namespace OfficeIMO.Yaml.Benchmarks.Workloads;

/// <summary>
/// VYaml in the parse-to-events and load lanes. VYaml reads UTF-8 bytes, so each input string is encoded once outside
/// timing and the byte array is reused, as the other libraries reuse their input string.
/// <list type="bullet">
/// <item>Parse: reads every event and materializes each scalar value, anchor name and tag as a string, as the other lanes do.</item>
/// <item>Load: VYaml has no representation tree; its dynamic deserialization builds dictionaries, lists and resolved
/// scalars for every document, which is the closest equivalent and is validated against YamlDotNet's loaded values.</item>
/// </list>
/// Load-and-save and emit are not offered: VYaml has no document model to save, and its emitter writes values rather
/// than replaying parser events with their anchors and tags.
/// </summary>
internal sealed class VYamlLibrary : IYamlReferenceLibrary {
    private readonly Dictionary<string, byte[]> _utf8 = new(ReferenceEqualityComparer.Instance);

    public static VYamlLibrary Instance { get; } = new();

    public string Name => "VYaml";

    public IYamlConformanceEngine Conformance { get; } = new VYamlConformanceEngine();

    public int ParseEvents(string yaml) {
        var parser = YamlParser.FromBytes(Utf8(yaml));
        parser.MaxDepth = VYamlConformanceEngine.MaxDepth;
        int count = 0;
        while (parser.Read()) {
            switch (parser.CurrentEventType) {
                case ParseEventType.Scalar:
                    _ = parser.GetScalarAsString();
                    if (parser.TryGetCurrentAnchor(out Anchor anchor)) _ = anchor.Name;
                    if (parser.TryGetCurrentTag(out Tag tag)) _ = VYamlConformanceEngine.ResolveTag(tag.Handle, tag.Suffix);
                    break;
                case ParseEventType.MappingStart:
                case ParseEventType.SequenceStart:
                    if (parser.TryGetCurrentAnchor(out Anchor collectionAnchor)) _ = collectionAnchor.Name;
                    if (parser.TryGetCurrentTag(out Tag collectionTag)) _ = VYamlConformanceEngine.ResolveTag(collectionTag.Handle, collectionTag.Suffix);
                    break;
                case ParseEventType.Alias:
                    if (parser.TryGetCurrentAnchor(out Anchor alias)) _ = alias.Name;
                    break;
            }

            count++;
        }

        return count;
    }

    public int LoadTree(string yaml) {
        var parser = YamlParser.FromBytes(Utf8(yaml));
        parser.MaxDepth = VYamlConformanceEngine.MaxDepth;
        int count = 0;
        foreach (object? document in YamlSerializer.DeserializeMultipleDocuments<object?>(ref parser)) count += Count(document);
        return count;
    }

    public string LoadAndSave(string yaml) => throw new NotSupportedException("VYaml has no document model to save.");

    public string EditAndSave(string yaml) => throw new NotSupportedException("VYaml has no document model to save.");

    public object CaptureEvents(string yaml) => throw new NotSupportedException("VYaml's emitter does not replay parser events.");

    public string Emit(object events) => throw new NotSupportedException("VYaml's emitter does not replay parser events.");

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
            case IDictionary<object, object?> mapping: {
                int count = 1;
                foreach (KeyValuePair<object, object?> entry in mapping) count += Count(entry.Key) + Count(entry.Value);
                return count;
            }
            case IList<object?> sequence: {
                int count = 1;
                foreach (object? item in sequence) count += Count(item);
                return count;
            }
            default:
                return 1;
        }
    }
}
