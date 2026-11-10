using OfficeIMO.Yaml.Benchmarks.Adapters;
using OfficeIMO.Yaml.Tests.Conformance;
using DotNet = YamlDotNet;
using Sharp = SharpYaml;

namespace OfficeIMO.Yaml.Benchmarks.Workloads;

/// <summary>The four measured operations, implemented directly against one reference library.</summary>
internal interface IYamlReferenceLibrary {
    string Name { get; }

    /// <summary>Adapter used only to validate outputs, outside timing.</summary>
    IYamlConformanceEngine Conformance { get; }

    /// <summary>Parses to events and returns the event count.</summary>
    int ParseEvents(string yaml);

    /// <summary>Loads the representation tree and returns its node count; an aliased subtree counts at each use.</summary>
    int LoadTree(string yaml);

    /// <summary>Loads the tree and saves it unchanged.</summary>
    string LoadAndSave(string yaml);

    /// <summary>Captures library-native events for <see cref="Emit"/>; runs outside timing.</summary>
    object CaptureEvents(string yaml);

    /// <summary>Emits captured events as a new document.</summary>
    string Emit(object events);
}

internal static class YamlReferenceLibraries {
    public static IYamlReferenceLibrary YamlDotNet { get; } = new YamlDotNetLibrary();

    public static IYamlReferenceLibrary SharpYaml { get; } = new SharpYamlLibrary();

    /// <summary>OfficeIMO.Yaml; emitting arrives with the writer.</summary>
    public static OfficeIMOYamlLibrary OfficeIMOYaml { get; } = new OfficeIMOYamlLibrary();

    internal sealed class OfficeIMOYamlLibrary : IYamlReferenceLibrary {
        public string Name => "OfficeIMO.Yaml";

        public IYamlConformanceEngine Conformance { get; } = new OfficeIMOYamlConformanceEngine();

        // Equivalent work to the reference parsers, whose events carry materialized values: every scalar value,
        // alias name, anchor and tag is produced as a string.
        public int ParseEvents(string yaml) {
            using var reader = new global::OfficeIMO.Yaml.YamlReader(yaml);
            int count = 0;
            while (reader.Read()) {
                switch (reader.TokenType) {
                    case global::OfficeIMO.Yaml.YamlTokenType.Scalar:
                        _ = reader.GetString();
                        _ = reader.Anchor;
                        _ = reader.Tag;
                        break;
                    case global::OfficeIMO.Yaml.YamlTokenType.MappingStart:
                    case global::OfficeIMO.Yaml.YamlTokenType.SequenceStart:
                        _ = reader.Anchor;
                        _ = reader.Tag;
                        break;
                    case global::OfficeIMO.Yaml.YamlTokenType.Alias:
                        _ = reader.GetString();
                        break;
                }

                count++;
            }

            return count;
        }

        // The model decodes scalar values on first use, so loading the tree leaves them undecoded; LoadTreeAndValues
        // measures the work of also reading every value as a string.
        public int LoadTree(string yaml) {
            global::OfficeIMO.Yaml.YamlDocument document = global::OfficeIMO.Yaml.YamlDocument.Parse(yaml);
            int count = 0;
            foreach (global::OfficeIMO.Yaml.YamlDocumentPart part in document.Documents) count += Count(part.Root, values: false);
            return count;
        }

        public int LoadTreeAndValues(string yaml) {
            global::OfficeIMO.Yaml.YamlDocument document = global::OfficeIMO.Yaml.YamlDocument.Parse(yaml);
            int count = 0;
            foreach (global::OfficeIMO.Yaml.YamlDocumentPart part in document.Documents) count += Count(part.Root, values: true);
            return count;
        }

        // Saving an unchanged document copies its retained source, so the saved text is the input string itself.
        public string LoadAndSave(string yaml) => global::OfficeIMO.Yaml.YamlDocument.Parse(yaml).ToString();

        public object CaptureEvents(string yaml) => throw new NotSupportedException("OfficeIMO.Yaml has no writer yet.");

        public string Emit(object events) => throw new NotSupportedException("OfficeIMO.Yaml has no writer yet.");

        private static int Count(global::OfficeIMO.Yaml.YamlNode node, bool values) {
            switch (node) {
                case global::OfficeIMO.Yaml.YamlAlias alias:
                    return Count(alias.Target, values);
                case global::OfficeIMO.Yaml.YamlSequence sequence: {
                    int count = 1;
                    foreach (global::OfficeIMO.Yaml.YamlNode item in sequence) count += Count(item, values);
                    return count;
                }
                case global::OfficeIMO.Yaml.YamlMapping mapping: {
                    int count = 1;
                    foreach (global::OfficeIMO.Yaml.YamlMappingEntry entry in mapping) count += Count(entry.Key, values) + Count(entry.Value, values);
                    return count;
                }
                default:
                    if (values) _ = node.AsString();
                    return 1;
            }
        }
    }

    private sealed class YamlDotNetLibrary : IYamlReferenceLibrary {
        public string Name => "YamlDotNet";

        public IYamlConformanceEngine Conformance { get; } = new YamlDotNetConformanceEngine();

        public int ParseEvents(string yaml) {
            var parser = new DotNet.Core.Parser(new StringReader(yaml));
            int count = 0;
            while (parser.MoveNext()) count++;
            return count;
        }

        public int LoadTree(string yaml) {
            var stream = new DotNet.RepresentationModel.YamlStream();
            stream.Load(new StringReader(yaml));
            int count = 0;
            foreach (DotNet.RepresentationModel.YamlDocument document in stream.Documents) count += Count(document.RootNode);
            return count;
        }

        public string LoadAndSave(string yaml) {
            var stream = new DotNet.RepresentationModel.YamlStream();
            stream.Load(new StringReader(yaml));
            var output = new StringWriter();
            stream.Save(output, assignAnchors: false);
            return output.ToString();
        }

        public object CaptureEvents(string yaml) {
            var parser = new DotNet.Core.Parser(new StringReader(yaml));
            var events = new List<DotNet.Core.Events.ParsingEvent>();
            while (parser.MoveNext()) events.Add(parser.Current!);
            return events;
        }

        public string Emit(object events) {
            var output = new StringWriter();
            var emitter = new DotNet.Core.Emitter(output);
            foreach (DotNet.Core.Events.ParsingEvent item in (List<DotNet.Core.Events.ParsingEvent>)events) emitter.Emit(item);
            return output.ToString();
        }

        private static int Count(DotNet.RepresentationModel.YamlNode node) {
            int count = 1;
            if (node is DotNet.RepresentationModel.YamlSequenceNode sequence) {
                foreach (DotNet.RepresentationModel.YamlNode child in sequence.Children) count += Count(child);
            } else if (node is DotNet.RepresentationModel.YamlMappingNode mapping) {
                foreach (KeyValuePair<DotNet.RepresentationModel.YamlNode, DotNet.RepresentationModel.YamlNode> entry in mapping.Children) count += Count(entry.Key) + Count(entry.Value);
            }

            return count;
        }
    }

    private sealed class SharpYamlLibrary : IYamlReferenceLibrary {
        public string Name => "SharpYaml";

        public IYamlConformanceEngine Conformance { get; } = new SharpYamlConformanceEngine();

        public int ParseEvents(string yaml) {
            Sharp.IParser parser = Sharp.Parser.CreateParser(new StringReader(yaml));
            int count = 0;
            while (parser.MoveNext()) count++;
            return count;
        }

        public int LoadTree(string yaml) {
            Sharp.Model.YamlStream stream = Sharp.Model.YamlStream.Load(new StringReader(yaml));
            int count = 0;
            foreach (Sharp.Model.YamlDocument document in stream) count += document.Contents == null ? 1 : Count(document.Contents);
            return count;
        }

        public string LoadAndSave(string yaml) {
            Sharp.Model.YamlStream stream = Sharp.Model.YamlStream.Load(new StringReader(yaml));
            var output = new StringWriter();
            stream.WriteTo(output, false);
            return output.ToString();
        }

        public object CaptureEvents(string yaml) {
            Sharp.IParser parser = Sharp.Parser.CreateParser(new StringReader(yaml));
            var events = new List<Sharp.Events.ParsingEvent>();
            while (parser.MoveNext()) events.Add(parser.Current!);
            return events;
        }

        public string Emit(object events) {
            var output = new StringWriter();
            var emitter = new Sharp.Emitter(output, 2, int.MaxValue, false, false, false);
            foreach (Sharp.Events.ParsingEvent item in (List<Sharp.Events.ParsingEvent>)events) emitter.Emit(item);
            return output.ToString();
        }

        private static int Count(Sharp.Model.YamlElement element) {
            int count = 1;
            if (element is Sharp.Model.YamlSequence sequence) {
                foreach (Sharp.Model.YamlElement child in sequence) count += Count(child);
            } else if (element is Sharp.Model.YamlMapping mapping) {
                foreach (KeyValuePair<Sharp.Model.YamlElement, Sharp.Model.YamlElement?> entry in mapping) count += Count(entry.Key) + (entry.Value == null ? 1 : Count(entry.Value));
            }

            return count;
        }
    }
}
