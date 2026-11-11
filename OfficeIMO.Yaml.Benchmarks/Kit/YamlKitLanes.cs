using System.Text;
using OfficeIMO.Yaml.Benchmarks.Adapters;
using OfficeIMO.Yaml.Tests.Conformance;
using One = OneYaml;
using Sharp = SharpYaml;

namespace OfficeIMO.Yaml.Benchmarks.Kit;

/// <summary>
/// One library's way of doing one kit task. <see cref="Prepare"/> builds the starting state outside timing (the input,
/// a value tree or a parsed document); <see cref="Run"/> is the timed work. An edit consumes its document, so the edit
/// lanes prepare a new state for every run. <see cref="Check"/> validates a result against the workload before timing.
/// </summary>
internal sealed class YamlKitLane {
    public YamlKitLane(string library, YamlKitTask task, Func<YamlKitWorkload, bool> applies, Func<YamlKitWorkload, object> prepare, Func<object, object?> run, Action<YamlKitWorkload, object?> check) {
        Library = library;
        Task = task;
        Applies = applies;
        Prepare = prepare;
        Run = run;
        Check = check;
    }

    public string Library { get; }

    public YamlKitTask Task { get; }

    public Func<YamlKitWorkload, bool> Applies { get; }

    public Func<YamlKitWorkload, object> Prepare { get; }

    public Func<object, object?> Run { get; }

    public Action<YamlKitWorkload, object?> Check { get; }

    /// <summary>Whether every run needs a state of its own because the run changes it.</summary>
    public bool ConsumesState => Task == YamlKitTask.EditScript;

    /// <summary>Prepares a state, runs it once and checks the result.</summary>
    public void Validate(YamlKitWorkload workload) {
        object state = Prepare(workload);
        Check(workload, Run(state));
        if (!ConsumesState) Check(workload, Run(state));
    }
}

/// <summary>
/// The kit's tasks for OfficeIMO.Yaml, OneYaml, VYaml and SharpYaml. Library columns:
/// <list type="bullet">
/// <item><c>OfficeIMO</c>: string in and out, the library's text API; typed loads use its generated context.</item>
/// <item><c>OfficeIMO reflection</c>: typed loads through reflection metadata.</item>
/// <item><c>OfficeIMO UTF-8</c>: UTF-8 bytes in (a stream over the input) and out (a stream, or for a dump the byte array OneYaml's dump returns), as the kit hands every library bytes.</item>
/// <item><c>OneYaml</c>: the kit's <c>oneyaml-csharp</c> row: reflection, UTF-8 bytes in and out, <see cref="One.YamlDocument"/> for documents.</item>
/// <item><c>OneYaml generated</c>: the kit's <c>oneyaml-csharp-gen</c> row: a generated context for typed and untyped values.</item>
/// <item><c>OneYaml string</c>: string in and out, the form the <c>OfficeIMO</c> column uses.</item>
/// <item><c>VYaml</c>: UTF-8 bytes and generated formatters.</item>
/// <item><c>SharpYaml</c>: its serializer with a generated context (typed) or <c>object</c> (untyped, references preserved), string in and out; its syntax tree for the document tasks.</item>
/// </list>
/// </summary>
internal static class YamlKitLanes {
    private static readonly UTF8Encoding Utf8 = new(false);
    private static readonly YamlSerializerOptions OfficeOptions = new() { PropertyNamingPolicy = YamlNamingPolicy.CamelCase };
    private static readonly KitYamlContext OfficeContext = new(new YamlSerializerOptions { PropertyNamingPolicy = YamlNamingPolicy.CamelCase });
    private static readonly KitOneYamlContext OneContext = KitOneYamlContext.Default;
    private static readonly Sharp.YamlSerializerOptions SharpLoad = new() { ReferenceHandling = Sharp.YamlReferenceHandling.Preserve };
    private static readonly Sharp.YamlSerializerOptions SharpDump = new();
    private static readonly Sharp.Syntax.YamlSyntaxOptions SharpSyntax = new() { IncludeTrivia = true };
    private static readonly YamlLoadOptions Unbounded = new() { MaxNodes = int.MaxValue };

    /// <summary>
    /// The OfficeIMO.Yaml document options: the defaults, which cap a document at 10,000,000 nodes, unless
    /// <c>--raise-limits</c> lifts the cap. The kit runs every library with its defaults.
    /// </summary>
    public static YamlLoadOptions? OfficeLoad { get; set; }

    private static readonly IYamlConformanceEngine DotNetReader = new YamlDotNetConformanceEngine();

    public static IReadOnlyList<string> Libraries { get; } = new[] {
        "OfficeIMO", "OfficeIMO reflection", "OfficeIMO UTF-8", "OneYaml", "OneYaml generated", "OneYaml string", "VYaml", "SharpYaml"
    };

    public static IReadOnlyList<YamlKitLane> All { get; } = Create();

    public static YamlKitLane? Find(string library, YamlKitTask task, YamlKitWorkload workload) =>
        All.FirstOrDefault(l => l.Library == library && l.Task == task && l.Applies(workload));

    private static List<YamlKitLane> Create() {
        var lanes = new List<YamlKitLane>();
        Func<YamlKitWorkload, bool> records = w => w.TypedModel == "records";
        Func<YamlKitWorkload, bool> gha = w => w.TypedModel == "gha";
        Func<YamlKitWorkload, bool> single = w => w.Documents == 1;
        Func<YamlKitWorkload, bool> any = _ => true;

        // Typed load.
        void Typed(string library, Func<YamlKitWorkload, object> prepare, Func<object, object?> readRecords, Func<object, object?> readWorkflow) {
            lanes.Add(new YamlKitLane(library, YamlKitTask.TypedLoad, records, prepare, readRecords, (w, r) => CheckTyped(w, r)));
            lanes.Add(new YamlKitLane(library, YamlKitTask.TypedLoad, gha, prepare, readWorkflow, (w, r) => CheckTyped(w, r)));
        }

        Typed("OfficeIMO", w => w.Text, s => YamlSerializer.Deserialize((string)s, OfficeContext.ListKitRecord), s => YamlSerializer.Deserialize((string)s, OfficeContext.KitWorkflow));
        Typed("OfficeIMO reflection", w => w.Text, s => YamlSerializer.Deserialize<List<KitRecord>>((string)s, OfficeOptions), s => YamlSerializer.Deserialize<KitWorkflow>((string)s, OfficeOptions));
        Typed("OfficeIMO UTF-8", w => w.Bytes, s => YamlSerializer.Deserialize(new MemoryStream((byte[])s, writable: false), OfficeContext.ListKitRecord), s => YamlSerializer.Deserialize(new MemoryStream((byte[])s, writable: false), OfficeContext.KitWorkflow));
        Typed("OneYaml", w => w.Bytes, s => One.YamlSerializer.Deserialize<List<KitRecord>>((byte[])s), s => One.YamlSerializer.Deserialize<KitWorkflow>((byte[])s));
        Typed("OneYaml generated", w => w.Bytes, s => One.YamlSerializer.Deserialize((byte[])s, OneContext.ListKitRecord), s => One.YamlSerializer.Deserialize((byte[])s, OneContext.KitWorkflow));
        Typed("OneYaml string", w => w.Text, s => One.YamlSerializer.Deserialize<List<KitRecord>>((string)s), s => One.YamlSerializer.Deserialize<KitWorkflow>((string)s));
        Typed("VYaml", w => w.Bytes, s => VYaml.Serialization.YamlSerializer.Deserialize<List<KitRecord>>((byte[])s), s => VYaml.Serialization.YamlSerializer.Deserialize<KitWorkflow>((byte[])s));
        Typed("SharpYaml", w => w.Text, s => Sharp.YamlSerializer.Deserialize((string)s, KitSharpYamlContext.Default.ListKitRecord), s => Sharp.YamlSerializer.Deserialize((string)s, KitSharpYamlContext.Default.KitWorkflow));

        // Generic load: every document of a stream.
        lanes.Add(new YamlKitLane("OfficeIMO", YamlKitTask.GenericLoad, any, w => w.Text,
            s => OfficeAll(new YamlReader((string)s)), CheckGeneric));
        lanes.Add(new YamlKitLane("OfficeIMO UTF-8", YamlKitTask.GenericLoad, any, w => w.Bytes,
            s => OfficeAll(new YamlReader(new MemoryStream((byte[])s, writable: false))), CheckGeneric));
        lanes.Add(new YamlKitLane("OneYaml", YamlKitTask.GenericLoad, any, w => w.Bytes,
            s => One.YamlSerializer.DeserializeAll<object>((byte[])s), CheckGeneric));
        lanes.Add(new YamlKitLane("OneYaml generated", YamlKitTask.GenericLoad, any, w => w.Bytes,
            s => One.YamlSerializer.DeserializeAll((byte[])s, OneContext.Object), CheckGeneric));
        lanes.Add(new YamlKitLane("OneYaml string", YamlKitTask.GenericLoad, any, w => w.Text,
            s => One.YamlSerializer.DeserializeAll<object>(new StringReader((string)s)), CheckGeneric));
        lanes.Add(new YamlKitLane("VYaml", YamlKitTask.GenericLoad, any, w => w.Bytes,
            s => VYaml.Serialization.YamlSerializer.DeserializeMultipleDocuments<object>((byte[])s).ToList(), CheckGeneric));
        lanes.Add(new YamlKitLane("SharpYaml", YamlKitTask.GenericLoad, single, w => w.Text,
            s => new object?[] { Sharp.YamlSerializer.Deserialize<object>((string)s, SharpLoad) }, CheckGeneric));

        // Dump: the library's own untyped tree, read before timing.
        lanes.Add(new YamlKitLane("OfficeIMO", YamlKitTask.Dump, single, w => YamlSerializer.Deserialize<object>(w.Text)!,
            s => YamlSerializer.Serialize<object>(s), CheckDump));
        lanes.Add(new YamlKitLane("OfficeIMO UTF-8", YamlKitTask.Dump, single, w => YamlSerializer.Deserialize<object>(w.Text)!,
            s => YamlSerializer.SerializeToUtf8Bytes<object>(s), CheckDump));
        lanes.Add(new YamlKitLane("OneYaml", YamlKitTask.Dump, single, w => One.YamlSerializer.Deserialize<object>(w.Bytes)!,
            s => One.YamlSerializer.SerializeToUtf8Bytes<object>(s), CheckDump));
        lanes.Add(new YamlKitLane("OneYaml generated", YamlKitTask.Dump, single, w => One.YamlSerializer.Deserialize(w.Bytes, OneContext.Object)!,
            s => One.YamlSerializer.SerializeToUtf8Bytes(s, OneContext.Object), CheckDump));
        lanes.Add(new YamlKitLane("OneYaml string", YamlKitTask.Dump, single, w => One.YamlSerializer.Deserialize<object>(w.Text)!,
            s => One.YamlSerializer.Serialize<object>(s), CheckDump));
        lanes.Add(new YamlKitLane("VYaml", YamlKitTask.Dump, single, w => VYaml.Serialization.YamlSerializer.Deserialize<object>(w.Bytes),
            s => VYaml.Serialization.YamlSerializer.Serialize<object?>(s), CheckDump));
        lanes.Add(new YamlKitLane("SharpYaml", YamlKitTask.Dump, single, w => Sharp.YamlSerializer.Deserialize<object>(w.Text, SharpLoad)!,
            s => Sharp.YamlSerializer.Serialize<object?>(s, SharpDump), CheckDump));

        // Document parse: the document is released after the run, as a caller's using block does.
        lanes.Add(new YamlKitLane("OfficeIMO", YamlKitTask.DocumentParse, any, w => w.Text, s => {
            using YamlDocument document = YamlDocument.Parse((string)s, OfficeLoad);
            return document.Documents.Count;
        }, (w, _) => CheckDocument(w, YamlDocument.Parse(w.Text, OfficeLoad).ToString())));
        lanes.Add(new YamlKitLane("OfficeIMO UTF-8", YamlKitTask.DocumentParse, any, w => w.Bytes, s => {
            using YamlDocument document = YamlDocument.Load(new MemoryStream((byte[])s, writable: false), OfficeLoad);
            return document.Documents.Count;
        }, (w, _) => CheckDocument(w, SaveUtf8(YamlDocument.Load(new MemoryStream(w.Bytes, writable: false), OfficeLoad)))));
        lanes.Add(new YamlKitLane("OneYaml", YamlKitTask.DocumentParse, any, w => w.Bytes, s => {
            using One.YamlDocument document = One.YamlDocument.Parse((byte[])s);
            return document.Documents.Count;
        }, (w, _) => CheckDocument(w, One.YamlDocument.Parse(w.Bytes).ToUtf8Bytes())));
        lanes.Add(new YamlKitLane("OneYaml string", YamlKitTask.DocumentParse, any, w => w.Text, s => {
            using One.YamlDocument document = One.YamlDocument.Parse((string)s);
            return document.Documents.Count;
        }, (w, _) => CheckDocument(w, One.YamlDocument.Parse(w.Text).ToString())));
        lanes.Add(new YamlKitLane("SharpYaml", YamlKitTask.DocumentParse, any, w => w.Text,
            s => Sharp.Syntax.YamlSyntaxTree.Parse((string)s, SharpSyntax), (w, r) => CheckBytes(w, w.Bytes, ((Sharp.Syntax.YamlSyntaxTree)r!).ToFullString())));

        // Unchanged emit: the document parsed before timing is written back.
        lanes.Add(new YamlKitLane("OfficeIMO", YamlKitTask.UnchangedEmit, any, w => YamlDocument.Parse(w.Text, OfficeLoad), s => ((YamlDocument)s).ToString(), (w, r) => CheckBytes(w, w.Bytes, r)));
        lanes.Add(new YamlKitLane("OfficeIMO UTF-8", YamlKitTask.UnchangedEmit, any, w => YamlDocument.Load(new MemoryStream(w.Bytes, writable: false), OfficeLoad), s => {
            var stream = new MemoryStream();
            ((YamlDocument)s).Save(stream);
            return stream;
        }, (w, r) => CheckBytes(w, w.Bytes, r)));
        lanes.Add(new YamlKitLane("OneYaml", YamlKitTask.UnchangedEmit, any, w => One.YamlDocument.Parse(w.Bytes), s => ((One.YamlDocument)s).ToUtf8Bytes(), (w, r) => CheckBytes(w, w.Bytes, r)));
        lanes.Add(new YamlKitLane("OneYaml string", YamlKitTask.UnchangedEmit, any, w => One.YamlDocument.Parse(w.Text), s => ((One.YamlDocument)s).ToString(), (w, r) => CheckBytes(w, w.Bytes, r)));
        lanes.Add(new YamlKitLane("SharpYaml", YamlKitTask.UnchangedEmit, any, w => Sharp.Syntax.YamlSyntaxTree.Parse(w.Text, SharpSyntax), s => ((Sharp.Syntax.YamlSyntaxTree)s).ToFullString(), (w, r) => CheckBytes(w, w.Bytes, r)));

        // Edit script: replace and remove through each library's own path API, then write the bytes.
        lanes.Add(new YamlKitLane("OfficeIMO", YamlKitTask.EditScript, any, w => new OfficeEdit(YamlDocument.Parse(w.Text, OfficeLoad), w), s => {
            var edit = (OfficeEdit)s;
            using YamlDocument document = edit.Apply();
            return document.ToString();
        }, (w, r) => CheckBytes(w, w.EditExpected!, r)));
        lanes.Add(new YamlKitLane("OfficeIMO UTF-8", YamlKitTask.EditScript, any, w => new OfficeEdit(YamlDocument.Load(new MemoryStream(w.Bytes, writable: false), OfficeLoad), w), s => {
            var edit = (OfficeEdit)s;
            using YamlDocument document = edit.Apply();
            var stream = new MemoryStream();
            document.Save(stream);
            return stream;
        }, (w, r) => CheckBytes(w, w.EditExpected!, r)));
        lanes.Add(new YamlKitLane("OneYaml", YamlKitTask.EditScript, any, w => new OneEdit(One.YamlDocument.Parse(w.Bytes), w), s => {
            using One.YamlDocument document = ((OneEdit)s).Apply();
            return document.ToUtf8Bytes();
        }, (w, r) => CheckBytes(w, w.EditExpected!, r)));
        lanes.Add(new YamlKitLane("OneYaml string", YamlKitTask.EditScript, any, w => new OneEdit(One.YamlDocument.Parse(w.Text), w), s => {
            using One.YamlDocument document = ((OneEdit)s).Apply();
            return document.ToString();
        }, (w, r) => CheckBytes(w, w.EditExpected!, r)));

        return lanes;
    }

    private static List<object?> OfficeAll(YamlReader reader) {
        using (reader) {
            var documents = new List<object?>();
            reader.Read();
            while (reader.TokenType != YamlTokenType.StreamEnd) {
                while (reader.TokenType is YamlTokenType.Comment or YamlTokenType.DocumentEnd or YamlTokenType.StreamStart) {
                    if (!reader.Read()) return documents;
                }

                if (reader.TokenType == YamlTokenType.StreamEnd) break;
                documents.Add(YamlSerializer.Deserialize<object>(reader));
            }

            return documents;
        }
    }

    private static MemoryStream SaveUtf8(YamlDocument document) {
        var stream = new MemoryStream();
        document.Save(stream);
        return stream;
    }

    private static void CheckTyped(YamlKitWorkload workload, object? result) {
        YamlConformanceNode node = result switch {
            List<KitRecord> list => YamlKitWorkload.ToNode(list),
            KitWorkflow workflow => YamlKitWorkload.ToNode(workflow),
            _ => throw new InvalidOperationException("Unexpected typed result " + result?.GetType().Name + ".")
        };
        Require(workload.Difference(new[] { node }), workload, "typed value");
    }

    private static void CheckGeneric(YamlKitWorkload workload, object? result) {
        var documents = ((System.Collections.IEnumerable)result!).Cast<object?>().Select(d => YamlKitWorkload.ToNode(d)).ToList();
        Require(workload.Difference(documents), workload, "loaded value");
    }

    // Two independent readers read the text back to the workload's value.
    private static void CheckDump(YamlKitWorkload workload, object? result) {
        byte[] bytes = ToBytes(result);
        YamlDocument read = YamlDocument.Load(new MemoryStream(bytes, writable: false), Unbounded);
        Require(workload.Difference(read.Documents.Select(part => OfficeIMOYamlConformanceEngine.ToConformance(part.Root)).ToList()), workload, "dump read by OfficeIMO.Yaml");
        Require(workload.Difference(DotNetReader.Load(bytes)), workload, "dump read by YamlDotNet");
    }

    private static void CheckDocument(YamlKitWorkload workload, object written) {
        CheckBytes(workload, workload.Bytes, written);
    }

    private static void CheckBytes(YamlKitWorkload workload, byte[] expected, object? result) {
        byte[] actual = ToBytes(result);
        if (!expected.AsSpan().SequenceEqual(actual)) {
            int index = 0;
            while (index < Math.Min(expected.Length, actual.Length) && expected[index] == actual[index]) index++;
            int start = Math.Max(0, index - 40);
            throw new InvalidOperationException(workload.Id + ": bytes differ at " + index + " of " + expected.Length + ": expected '"
                + Show(expected, start, index + 40) + "' but wrote '" + Show(actual, start, index + 40) + "'.");
        }
    }

    private static string Show(byte[] bytes, int start, int end) => Utf8.GetString(bytes, start, Math.Max(0, Math.Min(end, bytes.Length) - start)).Replace("\n", "\\n");

    private static byte[] ToBytes(object? result) => result switch {
        byte[] bytes => bytes,
        string text => Utf8.GetBytes(text),
        MemoryStream stream => stream.ToArray(),
        ReadOnlyMemory<byte> memory => memory.ToArray(),
        _ => throw new InvalidOperationException("Unexpected output " + result?.GetType().Name + ".")
    };

    private static void Require(string? difference, YamlKitWorkload workload, string what) {
        if (difference != null) throw new InvalidOperationException(workload.Id + ": " + what + " differs: " + difference);
    }

    /// <summary>A parsed OfficeIMO.Yaml document and the workload's edits, with paths resolved before timing.</summary>
    private sealed class OfficeEdit {
        private readonly YamlDocument _document;
        private readonly (bool Remove, string Path, YamlNode? Value)[] _steps;

        public OfficeEdit(YamlDocument document, YamlKitWorkload workload) {
            _document = document;
            _steps = workload.EditOps.Select(op => (op.Op == "remove", Path(document.Root, op.Tokens), Node(op.Value))).ToArray();
        }

        public YamlDocument Apply() {
            foreach ((bool remove, string path, YamlNode? value) in _steps) {
                if (remove) _document.Remove(path);
                else _document.Set(path, value);
            }

            return _document;
        }

        // A token is an index where the node it addresses is a sequence, as RFC 6901 reads it; keys are quoted.
        private static string Path(YamlNode? node, string[] tokens) {
            var path = new StringBuilder();
            foreach (string token in tokens) {
                if (node is YamlAlias alias) node = alias.Target;
                if (node is YamlSequence sequence) {
                    int index = int.Parse(token, System.Globalization.CultureInfo.InvariantCulture);
                    path.Append('[').Append(token).Append(']');
                    node = index < sequence.Count ? sequence[index] : null;
                } else {
                    path.Append("[\"").Append(token.Replace("\\", "\\\\").Replace("\"", "\\\"")).Append("\"]");
                    node = node is YamlMapping mapping && mapping.TryGetValue(token, out YamlNode? child) ? child : null;
                }
            }

            return path.ToString();
        }

        private static YamlNode? Node(object? value) => value switch {
            null => null,
            string text => text,
            bool flag => flag,
            long number => number,
            double number => number,
            _ => throw new NotSupportedException(value.GetType().Name)
        };
    }

    /// <summary>A parsed OneYaml document and the workload's edits; its pointer methods take RFC 6901 as written.</summary>
    private sealed class OneEdit {
        private readonly One.YamlDocument _document;
        private readonly YamlKitEdit[] _steps;

        public OneEdit(One.YamlDocument document, YamlKitWorkload workload) {
            _document = document;
            _steps = workload.EditOps.ToArray();
        }

        public One.YamlDocument Apply() {
            foreach (YamlKitEdit step in _steps) {
                if (step.Op == "remove") _document.Remove(step.Pointer);
                else _document.Set(step.Pointer, step.Value);
            }

            return _document;
        }
    }
}
